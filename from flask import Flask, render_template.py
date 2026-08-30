from flask import Flask, render_template, request, jsonify
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///mint_maintenance.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

# Database Models
class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    role = db.Column(db.String(20), default='CLIENT')  # CLIENT, SUPPLIER, ADMIN
    trade_category = db.Column(db.String(50), nullable=True)  # e.g., HVAC, Plumbing
    coverage_zip = db.Column(db.String(20), nullable=True)
    is_verified = db.Column(db.Boolean, default=False)

class Ticket(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    category = db.Column(db.String(50), nullable=False)
    description = db.Column(db.Text, nullable=False)
    address = db.Column(db.String(200), nullable=False)
    zip_code = db.Column(db.String(20), nullable=False)
    urgency = db.Column(db.String(20), default='STANDARD')
    status = db.Column(db.String(20), default='PENDING')  # PENDING, ASSIGNED, COMPLETED
    estimated_cost = db.Column(db.Float, nullable=False)
    client_name = db.Column(db.String(100), nullable=False)
    client_email = db.Column(db.String(120), nullable=False)
    supplier_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

BASE_RATES = {
    'HVAC': 120.0,
    'Plumbing': 95.0,
    'Electrical': 110.0,
    'Handyman': 75.0,
    'Renovation': 250.0
}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/admin')
def admin():
    tickets = Ticket.query.order_by(Ticket.created_at.desc()).all()
    suppliers = User.query.filter_by(role='SUPPLIER').all()
    return render_template('admin.html', tickets=tickets, suppliers=suppliers)

@app.route('/api/tickets', methods=['POST'])
def create_ticket():
    data = request.json
    category = data.get('category')
    urgency = data.get('urgency', 'STANDARD')

    # Pricing calculation algorithm
    base_rate = BASE_RATES.get(category, 100.0)
    multiplier = 1.5 if urgency == 'EMERGENCY_24_7' else (1.25 if urgency == 'SAME_DAY' else 1.0)
    estimated_cost = round(base_rate * multiplier, 2)

    new_ticket = Ticket(
        category=category,
        description=data.get('description'),
        address=data.get('address'),
        zip_code=data.get('zipCode'),
        urgency=urgency,
        estimated_cost=estimated_cost,
        client_name=data.get('clientName'),
        client_email=data.get('clientEmail')
    )

    # Automated supplier auto-matching by trade and zip code
    matched_supplier = User.query.filter_by(
        role='SUPPLIER',
        trade_category=category,
        coverage_zip=data.get('zipCode'),
        is_verified=True
    ).first()

    if matched_supplier:
        new_ticket.supplier_id = matched_supplier.id
        new_ticket.status = 'ASSIGNED'

    db.session.add(new_ticket)
    db.session.commit()

    return jsonify({
        'success': True,
        'ticket_id': new_ticket.id,
        'status': new_ticket.status,
        'estimated_cost': estimated_cost
    })

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True, port=5000)