"""
数据模型定义
包含产品、订单、订单明细、费用等表
"""
from sqlalchemy import Column, Integer, String, Float, Boolean, DateTime, ForeignKey, Text
from sqlalchemy.orm import relationship
from datetime import datetime
from database import Base


class Product(Base):
    """产品表"""
    __tablename__ = "products"

    id = Column(Integer, primary_key=True, index=True)
    model_code = Column(String(50), unique=True, nullable=False, index=True, comment="产品型号")
    product_name = Column(String(100), comment="产品名称")
    series_name = Column(String(50), comment="产品系列")
    cost_price = Column(Float, nullable=False, comment="成本价")
    default_price = Column(Float, comment="建议售价")
    is_active = Column(Boolean, default=True, comment="是否上架")
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    def __repr__(self):
        return f"<Product {self.model_code}>"


class Order(Base):
    """订单主表"""
    __tablename__ = "orders"

    id = Column(Integer, primary_key=True, index=True)
    order_no = Column(String(30), unique=True, nullable=False, index=True, comment="订单号")
    customer_name = Column(String(100), nullable=False, comment="客户名称")
    wood_frame_fee = Column(Float, default=0, comment="木架费")
    total_sales_amount = Column(Float, nullable=False, comment="订单总销售额")
    total_cost_amount = Column(Float, nullable=False, comment="订单总成本")
    profit = Column(Float, comment="利润")
    status = Column(String(20), default="ACTIVE", comment="状态: ACTIVE/CANCELLED")
    created_at = Column(DateTime, default=datetime.now)
    cancelled_at = Column(DateTime, nullable=True)

    # 关联
    items = relationship("OrderItem", back_populates="order", cascade="all, delete-orphan")

    def __repr__(self):
        return f"<Order {self.order_no}>"


class OrderItem(Base):
    """订单明细表"""
    __tablename__ = "order_items"

    id = Column(Integer, primary_key=True, index=True)
    order_id = Column(Integer, ForeignKey("orders.id"), nullable=False)
    product_model = Column(String(50), comment="产品型号")
    product_name = Column(String(100), comment="产品名称")
    series_name = Column(String(50), comment="产品系列")
    color_code = Column(String(50), comment="色号")
    quantity = Column(Integer, nullable=False, comment="数量")
    unit_price = Column(Float, nullable=False, comment="销售单价")
    cost_price = Column(Float, nullable=False, comment="成本单价")
    subtotal = Column(Float, nullable=False, comment="小计")

    # 关联
    order = relationship("Order", back_populates="items")

    def __repr__(self):
        return f"<OrderItem {self.product_model} x{self.quantity}>"


class Expense(Base):
    """费用表"""
    __tablename__ = "expenses"

    id = Column(Integer, primary_key=True, index=True)
    expense_date = Column(String(10), nullable=False, index=True, comment="费用日期(YYYY-MM-DD)")
    category = Column(String(50), nullable=False, comment="费用类别")
    amount = Column(Float, nullable=False, comment="金额")
    description = Column(Text, comment="备注")
    created_at = Column(DateTime, default=datetime.now)

    def __repr__(self):
        return f"<Expense {self.category} {self.amount}>"


class EmailConfig(Base):
    """邮件配置表"""
    __tablename__ = "email_config"

    id = Column(Integer, primary_key=True, index=True)
    smtp_server = Column(String(100), nullable=False, comment="SMTP服务器")
    smtp_port = Column(Integer, default=465, comment="SMTP端口")
    smtp_username = Column(String(100), nullable=False, comment="邮箱账号")
    smtp_password = Column(String(100), nullable=False, comment="邮箱密码")
    use_ssl = Column(Boolean, default=True, comment="是否使用SSL")
    sender_name = Column(String(50), default="涂料报价系统", comment="发件人名称")
    recipient_email = Column(String(100), nullable=False, comment="收件人邮箱")
    is_active = Column(Boolean, default=True, comment="是否启用")
    created_at = Column(DateTime, default=datetime.now)


class SystemConfig(Base):
    """系统配置表"""
    __tablename__ = "system_config"

    id = Column(Integer, primary_key=True, index=True)
    config_key = Column(String(50), unique=True, nullable=False, index=True)
    config_value = Column(Text, comment="配置值")
    description = Column(String(200), comment="说明")
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)


class FeishuConfig(Base):
    """飞书配置表"""
    __tablename__ = "feishu_config"

    id = Column(Integer, primary_key=True, index=True)
    webhook_url = Column(String(500), comment="Webhook地址")
    webhook_secret = Column(String(200), comment="签名密钥")
    keyword = Column(String(50), default="报价", comment="触发关键词")
    auto_send_report = Column(Boolean, default=False, comment="自动发送报表")
    report_time = Column(String(10), default="21:00", comment="报表发送时间")
    report_recipients = Column(String(500), comment="报表收件人(逗号分隔)")
    is_active = Column(Boolean, default=True, comment="是否启用")
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)
