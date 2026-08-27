# Loan Management System

A professional Flask-based loan management application with comprehensive features for financial tracking, client management, and reporting.

## Features

- 🔐 **Authentication System**: Admin and user roles with secure login
- 👥 **User Management**: Create and manage user accounts with specific permissions
- 📊 **Daily Collection Reports**: Track daily installments and payments
- 📈 **Sales Reports**: Comprehensive sales analysis and tracking
- 💰 **Payment Entry System**: Efficient payment processing and tracking
- 📋 **Loan Management**: Create, update, and close loans with business rule enforcement
- 📚 **Client Ledgers**: Individual and bulk client ledger views
- 💳 **Advances Reports**: Track advances and receivables
- 🚨 **Recovery System**: Track short payments and recovery actions
- 📄 **PDF/Excel Exports**: Professional export capabilities for all reports
- 📱 **PWA Support**: Progressive Web App for mobile access

## Tech Stack

- **Backend**: Flask, Flask-SQLAlchemy, Flask-Login
- **Database**: SQLite (development), PostgreSQL (production)
- **Frontend**: HTML5, Bootstrap 5, JavaScript
- **PDF Generation**: ReportLab, pdfkit
- **Excel Export**: XlsxWriter, pandas
- **Security**: Flask-WTF (CSRF protection), password hashing

## Installation

### Prerequisites
- Python 3.11+
- pip (Python package manager)

### Setup

1. **Clone the repository**
   ```bash
   git clone https://github.com/your-username/loanexter.git
   cd loanexter
   ```

2. **Create virtual environment**
   ```bash
   python -m venv .venv
   # Windows
   .venv\Scripts\activate
   # Linux/Mac
   source .venv/bin/activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Configure environment variables**
   ```bash
   cp .env.example .env
   # Edit .env with your configuration
   ```

5. **Run the application**
   ```bash
   python app.py
   ```

The application will be available at `http://127.0.0.1:5000`

## Default Admin Credentials

- **Username**: `shivam`
- **Password**: `Raaina@20`

⚠️ **Important**: Change these credentials in production using environment variables:
```env
ADMIN_USERNAME=your_secure_username
ADMIN_PASSWORD=your_secure_password
```

## Environment Variables

Create a `.env` file with the following variables:

```env
# Database Configuration
DATABASE_URL=postgresql://username:password@localhost:5432/loanexter_db

# Admin User Configuration
ADMIN_USERNAME=shivam
ADMIN_PASSWORD=Raaina@20

# Flask Configuration
SECRET_KEY=your-secret-key-here
FLASK_DEBUG=False
```

## Project Structure

```
loanexter/
├── app.py                 # Main Flask application
├── requirements.txt        # Python dependencies
├── .env.example          # Environment variables template
├── .gitignore            # Git ignore rules
├── templates/            # HTML templates
│   ├── index.html       # Main dashboard
│   ├── login.html       # Login page
│   ├── admin_users.html # User management
│   └── ...              # Other templates
├── static/              # Static files (CSS, JS, images)
└── README.md            # This file
```

## Security Features

- 🔐 Password hashing with Werkzeug
- 🛡️ CSRF protection with Flask-WTF
- 🔒 Session security with HTTPOnly cookies
- 👤 Role-based access control
- 📋 Page-level permissions
- 🚫 SQL injection prevention with SQLAlchemy ORM

## Deployment

### Production Deployment

1. **Set up PostgreSQL database**
2. **Configure environment variables** on your hosting platform
3. **Install dependencies** on your server
4. **Run with production WSGI server** (gunicorn)

Example deployment command:
```bash
gunicorn -w 4 -b 0.0.0.0:5000 app:app
```

### Recommended Hosting Platforms

- **Render**: Easy Flask deployment with PostgreSQL
- **Heroku**: Popular PaaS with good Flask support
- **Railway**: Modern platform with database support
- **DigitalOcean**: Full control with droplets

## Business Rules Implemented

- **Daily Loans**: 100 installments, starts from next day
- **Weekly Loans**: 14 equal installments + 2 days balance
- **Ten-Day Loans**: 10 installments, every 10 days
- **Loan Closing**: Close today → installment = 0 tomorrow
- **Weekly Advance**: First week free, Day 98 = ₹400 only

## Support

For issues and questions, please create an issue in the GitHub repository.

## License

This project is proprietary software. All rights reserved.