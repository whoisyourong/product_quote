from datetime import datetime

from flask_sqlalchemy import SQLAlchemy


db = SQLAlchemy()


class Product(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    product_code = db.Column(db.String(100), nullable=False, unique=True, index=True)
    product_name = db.Column(db.String(100), nullable=False)
    specification = db.Column(db.String(200))
    smt_points = db.Column(db.Integer, default=0)
    welding_points = db.Column(db.Integer, default=0)
    binding_wires = db.Column(db.Integer, default=0)
    labor_cost = db.Column(db.Float, default=0)
    packaging_cost = db.Column(db.Float, default=0)
    imported_at = db.Column(db.DateTime, default=datetime.utcnow)

    bom_items = db.relationship(
        "ProductMaterial",
        back_populates="product",
        cascade="all, delete-orphan",
        lazy="selectin",
    )

    def __repr__(self):
        return f"<Product {self.product_code} - {self.product_name}>"


class Material(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    part_code = db.Column(db.String(100), nullable=False, unique=True, index=True)
    part_name = db.Column(db.String(100), nullable=False)
    specification = db.Column(db.String(200))
    unit = db.Column(db.String(50))
    price = db.Column(db.Float, default=0)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow)

    bom_items = db.relationship("ProductMaterial", back_populates="material")

    def __repr__(self):
        return f"<Material {self.part_code} - {self.part_name}>"


class ProductMaterial(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    product_id = db.Column(db.Integer, db.ForeignKey("product.id"), nullable=False)
    material_id = db.Column(db.Integer, db.ForeignKey("material.id"), nullable=False)
    quantity = db.Column(db.Float, nullable=False, default=0)

    product = db.relationship("Product", back_populates="bom_items")
    material = db.relationship("Material", back_populates="bom_items")


class ProcessRate(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    rate_key = db.Column(db.String(50), nullable=False, unique=True, index=True)
    rate_name = db.Column(db.String(100), nullable=False)
    unit_price = db.Column(db.Float, default=0)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow)


class Quote(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    quote_type = db.Column(db.String(30), default="bom")
    quote_date = db.Column(db.DateTime, default=datetime.utcnow)
    customer_name = db.Column(db.String(100), nullable=False)
    product_id = db.Column(db.Integer)
    product_code = db.Column(db.String(100))
    product_name = db.Column(db.String(100))
    specification = db.Column(db.String(200))
    material_snapshot = db.Column(db.Text)
    product_snapshot = db.Column(db.Text)
    material_cost = db.Column(db.Float, default=0)
    process_cost = db.Column(db.Float, default=0)
    calculated_cost = db.Column(db.Float, default=0)
    markup_percentage = db.Column(db.Float, default=0)
    final_price = db.Column(db.Float, default=0)
    note = db.Column(db.Text)
    deleted_at = db.Column(db.DateTime)

    def __repr__(self):
        return f"<Quote {self.id} - {self.customer_name}>"
