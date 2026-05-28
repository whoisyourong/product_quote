from datetime import datetime
import json
import re
import zipfile
import xml.etree.ElementTree as ET

from flask import Flask, flash, jsonify, redirect, render_template, request, url_for
from sqlalchemy import inspect, text

from models import Material, ProcessRate, Product, ProductMaterial, Quote, db


app = Flask(__name__)
app.config["SECRET_KEY"] = "dev-secret-key"
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///product_quote_system.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db.init_app(app)


NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def round_half_up(value, digits=2):
    from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

    try:
        number = Decimal(str(value or 0))
    except (InvalidOperation, ValueError):
        number = Decimal("0")
    quant = Decimal("1") if digits == 0 else Decimal("1").scaleb(-digits)
    return number.quantize(quant, rounding=ROUND_HALF_UP)


def format_money(value):
    return f"{round_half_up(value, 2):.2f}"


def format_material_price(value):
    return f"{round_half_up(value, 4):.4f}"


def format_quantity(value):
    return str(int(round_half_up(value, 0)))


app.jinja_env.filters["money"] = format_money
app.jinja_env.filters["material_price"] = format_material_price
app.jinja_env.filters["quantity"] = format_quantity


def ensure_schema():
    db.create_all()
    inspector = inspect(db.engine)
    table_columns = {
        table: {column["name"] for column in inspector.get_columns(table)}
        for table in inspector.get_table_names()
    }

    statements = []
    product_columns = table_columns.get("product", set())
    if "imported_at" not in product_columns:
        statements.append("ALTER TABLE product ADD COLUMN imported_at DATETIME")

    material_columns = table_columns.get("material", set())
    if "updated_at" not in material_columns:
        statements.append("ALTER TABLE material ADD COLUMN updated_at DATETIME")

    quote_columns = table_columns.get("quote", set())
    quote_additions = {
        "quote_type": "VARCHAR(30)",
        "product_id": "INTEGER",
        "product_code": "VARCHAR(100)",
        "product_name": "VARCHAR(100)",
        "specification": "VARCHAR(200)",
        "product_snapshot": "TEXT",
        "material_cost": "FLOAT",
        "process_cost": "FLOAT",
        "profit_margin_percentage": "FLOAT",
        "note": "TEXT",
        "deleted_at": "DATETIME",
    }
    for column, column_type in quote_additions.items():
        if column not in quote_columns:
            statements.append(f"ALTER TABLE quote ADD COLUMN {column} {column_type}")

    for statement in statements:
        db.session.execute(text(statement))

    default_rates = {
        "smt": "贴片单价",
        "welding": "焊接单价",
        "binding": "绑定单价",
    }
    for rate_key, rate_name in default_rates.items():
        rate = ProcessRate.query.filter_by(rate_key=rate_key).first()
        if not rate:
            db.session.add(ProcessRate(rate_key=rate_key, rate_name=rate_name, unit_price=0))
    db.session.commit()


with app.app_context():
    ensure_schema()


def parse_number(value, default=0):
    if value in (None, ""):
        return default
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def parse_integer(value, default=0):
    return int(round_half_up(parse_number(value, default), 0))


def calculate_profit_margin(final_price, total_cost):
    if final_price <= 0:
        return 0
    return (final_price - total_cost) / final_price * 100


def quote_profit_margin(quote):
    stored_margin = getattr(quote, "profit_margin_percentage", None)
    if stored_margin is not None:
        return stored_margin
    return calculate_profit_margin(quote.final_price or 0, quote.calculated_cost or 0)


app.jinja_env.globals["quote_profit_margin"] = quote_profit_margin


def split_label(value, label):
    text_value = str(value or "").strip()
    return re.sub(rf"^{label}\s*[:：]\s*", "", text_value).strip()


def read_xlsx_rows(file_storage):
    content = file_storage.read()
    with zipfile.ZipFile(__import__("io").BytesIO(content)) as zf:
        shared_strings = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for item in root.findall("a:si", NS):
                shared_strings.append("".join(t.text or "" for t in item.findall(".//a:t", NS)))

        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        first_sheet = workbook.find("a:sheets/a:sheet", NS)
        rel_id = first_sheet.attrib[f"{{{NS['r']}}}id"]
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        targets = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        target = targets[rel_id].lstrip("/")
        if not target.startswith("xl/"):
            target = f"xl/{target}"

        sheet = ET.fromstring(zf.read(target))
        rows = []
        for row in sheet.findall("a:sheetData/a:row", NS):
            values = []
            for cell in row.findall("a:c", NS):
                cell_type = cell.attrib.get("t")
                value_node = cell.find("a:v", NS)
                value = "" if value_node is None else value_node.text or ""
                if cell_type == "s" and value:
                    value = shared_strings[int(value)]
                elif cell_type == "inlineStr":
                    value = "".join(t.text or "" for t in cell.findall(".//a:t", NS))
                values.append(value)
            rows.append(values)
        return rows


def import_bom(file_storage):
    rows = read_xlsx_rows(file_storage)
    if len(rows) < 3:
        raise ValueError("Excel 内容太少，未找到产品和材料明细。")

    header = rows[0]
    product_code = split_label(header[0] if len(header) > 0 else "", "产品编码")
    product_name = split_label(header[1] if len(header) > 1 else "", "品名")
    specification = split_label(header[2] if len(header) > 2 else "", "规格")
    if not product_code or not product_name:
        raise ValueError("第 1 行需要包含产品编码和品名。")

    product = Product.query.filter_by(product_code=product_code).first()
    if product:
        raise ValueError(f"产品编码 {product_code} 已经导入过，请勿重复导入。")

    product = Product(product_code=product_code, product_name=product_name)
    db.session.add(product)

    product.product_name = product_name
    product.specification = specification
    product.imported_at = datetime.utcnow()
    product.bom_items.clear()

    imported_count = 0
    for row in rows[2:]:
        if len(row) < 5:
            continue
        part_code = str(row[0] or "").strip()
        part_name = str(row[1] or "").strip()
        if not part_code or not part_name:
            continue

        material = Material.query.filter_by(part_code=part_code).first()
        if not material:
            material = Material(part_code=part_code, price=0)
            db.session.add(material)

        material.part_name = part_name
        material.specification = str(row[2] or "").strip()
        material.unit = str(row[3] or "").strip()

        product.bom_items.append(
            ProductMaterial(material=material, quantity=parse_integer(row[4]))
        )
        imported_count += 1

    db.session.commit()
    return product, imported_count


def get_process_rates():
    rates = {rate.rate_key: rate for rate in ProcessRate.query.all()}
    for rate_key, rate_name in {
        "smt": "贴片单价",
        "welding": "焊接单价",
        "binding": "绑定单价",
    }.items():
        if rate_key not in rates:
            rates[rate_key] = ProcessRate(rate_key=rate_key, rate_name=rate_name, unit_price=0)
            db.session.add(rates[rate_key])
    db.session.commit()
    return rates


def product_cost(product):
    material_rows = []
    raw_material_cost = 0
    for item in product.bom_items:
        line_cost = (item.quantity or 0) * (item.material.price or 0)
        raw_material_cost += line_cost
        material_rows.append(
            {
                "id": item.id,
                "part_code": item.material.part_code,
                "part_name": item.material.part_name,
                "specification": item.material.specification,
                "unit": item.material.unit,
                "quantity": item.quantity or 0,
                "price": item.material.price or 0,
                "line_cost": line_cost,
            }
        )

    material_loss_cost = raw_material_cost * 0.01
    material_cost = raw_material_cost + material_loss_cost
    rates = get_process_rates()
    smt_cost = (product.smt_points or 0) * (rates["smt"].unit_price or 0)
    welding_cost = (product.welding_points or 0) * (rates["welding"].unit_price or 0)
    binding_cost = (product.binding_wires or 0) * (rates["binding"].unit_price or 0)
    process_detail = {
        "raw_material_cost": raw_material_cost,
        "material_loss_rate": 0.01,
        "material_loss_cost": material_loss_cost,
        "smt_points": product.smt_points or 0,
        "smt_unit_price": rates["smt"].unit_price or 0,
        "smt_cost": smt_cost,
        "welding_points": product.welding_points or 0,
        "welding_unit_price": rates["welding"].unit_price or 0,
        "welding_cost": welding_cost,
        "binding_wires": product.binding_wires or 0,
        "binding_unit_price": rates["binding"].unit_price or 0,
        "binding_cost": binding_cost,
        "labor_cost": product.labor_cost or 0,
        "packaging_cost": product.packaging_cost or 0,
    }
    process_cost = (
        smt_cost
        + welding_cost
        + binding_cost
        + (product.labor_cost or 0)
        + (product.packaging_cost or 0)
    )
    total_cost = material_cost + process_cost
    return material_cost, process_cost, total_cost, material_rows, process_detail


def calculate_process_cost(raw_material_cost, smt_points, welding_points, binding_wires, labor_cost, packaging_cost):
    material_loss_cost = raw_material_cost * 0.01
    material_cost = raw_material_cost + material_loss_cost
    rates = get_process_rates()
    smt_cost = smt_points * (rates["smt"].unit_price or 0)
    welding_cost = welding_points * (rates["welding"].unit_price or 0)
    binding_cost = binding_wires * (rates["binding"].unit_price or 0)
    process_cost = smt_cost + welding_cost + binding_cost + labor_cost + packaging_cost
    process_detail = {
        "raw_material_cost": raw_material_cost,
        "material_loss_rate": 0.01,
        "material_loss_cost": material_loss_cost,
        "smt_points": smt_points,
        "smt_unit_price": rates["smt"].unit_price or 0,
        "smt_cost": smt_cost,
        "welding_points": welding_points,
        "welding_unit_price": rates["welding"].unit_price or 0,
        "welding_cost": welding_cost,
        "binding_wires": binding_wires,
        "binding_unit_price": rates["binding"].unit_price or 0,
        "binding_cost": binding_cost,
        "labor_cost": labor_cost,
        "packaging_cost": packaging_cost,
    }
    return material_cost, process_cost, material_cost + process_cost, process_detail


def next_temporary_product_code():
    today_prefix = f"TEMP-{datetime.utcnow().strftime('%Y%m%d')}-"
    today_count = Quote.query.filter(
        Quote.quote_type == "temporary",
        Quote.product_code.like(f"{today_prefix}%"),
    ).count()
    return f"{today_prefix}{today_count + 1:02d}"


@app.route("/")
def index():
    products = Product.query.order_by(Product.imported_at.desc()).limit(6).all()
    quote_count = Quote.query.filter(Quote.deleted_at.is_(None)).count()
    material_count = Material.query.count()
    product_count = Product.query.count()
    recent_quotes = (
        Quote.query.filter(Quote.deleted_at.is_(None))
        .order_by(Quote.quote_date.desc())
        .limit(5)
        .all()
    )
    return render_template(
        "index.html",
        products=products,
        product_count=product_count,
        material_count=material_count,
        quote_count=quote_count,
        recent_quotes=recent_quotes,
    )


@app.route("/import", methods=["POST"])
def import_excel():
    excel_file = request.files.get("excel_file")
    if not excel_file or not excel_file.filename:
        flash("请选择一个 Excel 文件。", "error")
        return redirect(url_for("index"))

    try:
        product, imported_count = import_bom(excel_file)
    except Exception as exc:
        flash(f"导入失败：{exc}", "error")
        return redirect(url_for("index"))

    flash(f"已导入 {product.product_code}，材料明细 {imported_count} 条。", "success")
    return redirect(url_for("product_detail", product_id=product.id))


@app.route("/products")
def products():
    product_list = Product.query.order_by(Product.product_code.asc()).all()
    costs = {product.id: product_cost(product)[:3] for product in product_list}
    return render_template("products.html", products=product_list, costs=costs)


@app.route("/products/<int:product_id>", methods=["GET", "POST"])
def product_detail(product_id):
    product = Product.query.get_or_404(product_id)
    if request.method == "POST":
        product.smt_points = parse_integer(request.form.get("smt_points"))
        product.welding_points = parse_integer(request.form.get("welding_points"))
        product.binding_wires = parse_integer(request.form.get("binding_wires"))
        product.labor_cost = parse_number(request.form.get("labor_cost"))
        product.packaging_cost = parse_number(request.form.get("packaging_cost"))
        db.session.commit()
        flash("产品工艺费用已更新。", "success")
        return redirect(url_for("product_detail", product_id=product.id))

    material_cost, process_cost, total_cost, material_rows, process_detail = product_cost(product)
    return render_template(
        "product_detail.html",
        product=product,
        material_cost=material_cost,
        process_cost=process_cost,
        total_cost=total_cost,
        material_rows=material_rows,
        process_detail=process_detail,
    )


@app.route("/products/<int:product_id>/bom-items/<int:item_id>/delete", methods=["POST"])
def delete_product_bom_item(product_id, item_id):
    product = Product.query.get_or_404(product_id)
    item = ProductMaterial.query.filter_by(id=item_id, product_id=product.id).first_or_404()
    material_name = item.material.part_name if item.material else ""
    db.session.delete(item)
    db.session.commit()
    flash(f"BOM 材料 {material_name} 已删除。", "success")
    return redirect(url_for("product_detail", product_id=product.id))


@app.route("/materials", methods=["GET", "POST"])
def materials():
    if request.method == "POST":
        for material in Material.query.all():
            value = request.form.get(f"price_{material.id}")
            if value is not None:
                material.price = parse_number(value)
                material.updated_at = datetime.utcnow()
        db.session.commit()
        flash("材料价格已更新，产品成本会实时刷新。", "success")
        return redirect(url_for("materials"))

    material_list = Material.query.order_by(Material.part_code.asc()).all()
    return render_template("materials.html", materials=material_list)


@app.route("/materials/<int:material_id>/price", methods=["POST"])
def update_material_price(material_id):
    material = Material.query.get_or_404(material_id)
    data = request.get_json(silent=True) or {}
    material.price = parse_number(data.get("price"))
    material.updated_at = datetime.utcnow()
    db.session.commit()
    return jsonify(
        {
            "ok": True,
            "price": material.price,
            "updated_at": material.updated_at.strftime("%Y-%m-%d %H:%M"),
        }
    )


@app.route("/materials/<int:material_id>/delete", methods=["POST"])
def delete_material(material_id):
    material = Material.query.get_or_404(material_id)
    material_name = material.part_name
    linked_count = ProductMaterial.query.filter_by(material_id=material.id).count()
    ProductMaterial.query.filter_by(material_id=material.id).delete(synchronize_session=False)
    db.session.delete(material)
    db.session.commit()
    if linked_count:
        flash(f"材料 {material_name} 已删除，并同步移除了 {linked_count} 条产品 BOM 引用。", "success")
    else:
        flash(f"材料 {material_name} 已删除。", "success")
    return redirect(url_for("materials"))


@app.route("/process-rates")
def process_rates():
    rates = ProcessRate.query.order_by(ProcessRate.id.asc()).all()
    return render_template("process_rates.html", rates=rates)


@app.route("/process-rates/<int:rate_id>/price", methods=["POST"])
def update_process_rate(rate_id):
    rate = ProcessRate.query.get_or_404(rate_id)
    data = request.get_json(silent=True) or {}
    rate.unit_price = parse_number(data.get("unit_price"))
    rate.updated_at = datetime.utcnow()
    db.session.commit()
    return jsonify(
        {
            "ok": True,
            "unit_price": rate.unit_price,
            "updated_at": rate.updated_at.strftime("%Y-%m-%d %H:%M"),
        }
    )


@app.route("/temporary-quote", methods=["GET", "POST"])
def temporary_quote():
    materials = Material.query.order_by(Material.part_code.asc()).all()
    process_rates = get_process_rates()
    rate_values = {
        "smt": process_rates["smt"].unit_price or 0,
        "welding": process_rates["welding"].unit_price or 0,
        "binding": process_rates["binding"].unit_price or 0,
    }
    material_options = [
        {
            "id": material.id,
            "part_code": material.part_code,
            "part_name": material.part_name,
            "specification": material.specification or "",
            "unit": material.unit or "",
            "price": material.price or 0,
        }
        for material in materials
    ]
    if request.method == "GET":
        return render_template(
            "temporary_quote.html",
            materials=materials,
            material_options=material_options,
            rate_values=rate_values,
        )

    row_modes = request.form.getlist("row_mode")
    material_ids = request.form.getlist("material_id")
    part_codes = request.form.getlist("part_code")
    part_names = request.form.getlist("part_name")
    specifications = request.form.getlist("material_specification")
    units = request.form.getlist("unit")
    quantities = request.form.getlist("quantity")
    prices = request.form.getlist("price")

    material_rows = []
    raw_material_cost = 0
    for index, row_mode in enumerate(row_modes):
        quantity = parse_integer(quantities[index] if index < len(quantities) else 0)
        price = parse_number(prices[index] if index < len(prices) else 0)
        if quantity <= 0:
            continue

        material = None
        if row_mode == "existing" and index < len(material_ids) and material_ids[index]:
            material = Material.query.get(material_ids[index])

        part_code = material.part_code if material else (part_codes[index] if index < len(part_codes) else "")
        part_name = material.part_name if material else (part_names[index] if index < len(part_names) else "")
        specification = material.specification if material else (specifications[index] if index < len(specifications) else "")
        unit = material.unit if material else (units[index] if index < len(units) else "")
        if not part_name:
            continue

        line_cost = quantity * price
        raw_material_cost += line_cost
        material_rows.append(
            {
                "source": "existing" if material else "temporary",
                "part_code": part_code,
                "part_name": part_name,
                "specification": specification,
                "unit": unit,
                "quantity": quantity,
                "price": price,
                "line_cost": line_cost,
            }
        )

    if not material_rows:
        flash("请至少录入一条有效材料。", "error")
        return render_template(
            "temporary_quote.html",
            materials=materials,
            material_options=material_options,
            rate_values=rate_values,
        )

    smt_points = parse_integer(request.form.get("smt_points"))
    welding_points = parse_integer(request.form.get("welding_points"))
    binding_wires = parse_integer(request.form.get("binding_wires"))
    labor_cost = parse_number(request.form.get("labor_cost"))
    packaging_cost = parse_number(request.form.get("packaging_cost"))
    material_cost, process_cost, total_cost, process_detail = calculate_process_cost(
        raw_material_cost,
        smt_points,
        welding_points,
        binding_wires,
        labor_cost,
        packaging_cost,
    )
    final_price = parse_number(request.form.get("final_price"))
    if final_price <= 0:
        flash("最终报价金额必须大于 0。", "error")
        return render_template(
            "temporary_quote.html",
            materials=materials,
            material_options=material_options,
            rate_values=rate_values,
        )
    profit_margin = calculate_profit_margin(final_price, total_cost)
    product_code = request.form.get("product_code") or next_temporary_product_code()

    quote = Quote(
        quote_type="temporary",
        customer_name=request.form["customer_name"],
        product_code=product_code,
        product_name=request.form.get("product_name") or "临时报价产品",
        specification=request.form.get("product_specification"),
        material_snapshot=json.dumps(material_rows, ensure_ascii=False),
        product_snapshot=json.dumps(process_detail, ensure_ascii=False),
        material_cost=material_cost,
        process_cost=process_cost,
        calculated_cost=total_cost,
        markup_percentage=0,
        profit_margin_percentage=profit_margin,
        final_price=final_price,
        note=request.form.get("note"),
        quote_date=datetime.utcnow(),
    )
    db.session.add(quote)
    db.session.commit()
    flash("临时报价已保存为不可变快照。", "success")
    return redirect(url_for("quote_detail", quote_id=quote.id))


@app.route("/products/<int:product_id>/quote", methods=["POST"])
def create_quote(product_id):
    product = Product.query.get_or_404(product_id)
    material_cost, process_cost, total_cost, material_rows, process_detail = product_cost(product)
    final_price = parse_number(request.form.get("final_price"))
    if final_price <= 0:
        flash("最终报价金额必须大于 0。", "error")
        return redirect(url_for("product_detail", product_id=product.id))
    profit_margin = calculate_profit_margin(final_price, total_cost)

    quote = Quote(
        quote_type="bom",
        customer_name=request.form["customer_name"],
        product_id=product.id,
        product_code=product.product_code,
        product_name=product.product_name,
        specification=product.specification,
        material_snapshot=json.dumps(material_rows, ensure_ascii=False),
        product_snapshot=json.dumps(
            {
                **process_detail,
            },
            ensure_ascii=False,
        ),
        material_cost=material_cost,
        process_cost=process_cost,
        calculated_cost=total_cost,
        markup_percentage=0,
        profit_margin_percentage=profit_margin,
        final_price=final_price,
        note=request.form.get("note"),
        quote_date=datetime.utcnow(),
    )
    db.session.add(quote)
    db.session.commit()
    flash("报价已保存为不可变快照。", "success")
    return redirect(url_for("quote_detail", quote_id=quote.id))


@app.route("/quotes")
def quotes():
    quote_list = (
        Quote.query.filter(Quote.deleted_at.is_(None))
        .order_by(Quote.quote_date.desc())
        .all()
    )
    return render_template("quotes.html", quotes=quote_list)


@app.route("/quotes/<int:quote_id>")
def quote_detail(quote_id):
    quote = Quote.query.filter(
        Quote.id == quote_id,
        Quote.deleted_at.is_(None),
    ).first_or_404()
    materials_snapshot = json.loads(quote.material_snapshot or "[]")
    product_snapshot = json.loads(quote.product_snapshot or "{}")
    return render_template(
        "quote_detail.html",
        quote=quote,
        materials_snapshot=materials_snapshot,
        product_snapshot=product_snapshot,
    )


@app.route("/quotes/<int:quote_id>/delete", methods=["POST"])
def delete_quote(quote_id):
    quote = Quote.query.filter(
        Quote.id == quote_id,
        Quote.deleted_at.is_(None),
    ).first_or_404()
    quote.deleted_at = datetime.utcnow()
    db.session.commit()
    flash(f"报价 #{quote.id} 已从档案中删除。", "success")
    return redirect(url_for("quotes"))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=False)
