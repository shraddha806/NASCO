from flask import Flask, request, send_file
from flask_cors import CORS
from flask_restx import Api, Resource, fields, Namespace
import asyncio
import sys
import os
import json
from datetime import datetime
from dotenv import load_dotenv

# Gmail inbox watcher (in-memory, no MongoDB)
from notifications import (
    get_missed_email_by_id,
    get_claims_dashboard,
    start_inbox_watcher,
    GMAIL_USER,
    INBOX_POLL_SEC,
    _inbox_watcher_started,
)

# Fix Windows asyncio subprocess issues
if sys.platform == 'win32':
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

load_dotenv()

# Import automation functions
from pro import run_icici_lombard, run_policybazaar, run_insurancedekho
from pro import VEHICLE_REG_NO, VALID_USERNAME, EMAIL_ID, FULL_NAME

app = Flask(__name__)

# Configure CORS for both development and production
cors_origins = [
    "http://localhost:3000",           # Development
    "https://www.insurebot.com",       # Production frontend
    "https://insurebot.com",           # Production (without www)
]
CORS(app, origins=cors_origins, supports_credentials=True)

# Initialize Flask-RESTX API with Swagger documentation
api = Api(
    app,
    version='1.0.0',
    title='Insurance Quote API',
    description='REST API for fetching and comparing insurance quotes from multiple providers',
    doc='/swagger',  # Swagger UI endpoint
    prefix='/api',
    authorizations={
        'apikey': {
            'type': 'apiKey',
            'in': 'header',
            'name': 'X-API-Key'
        }
    }
)

# Create namespaces for API organization
ns_health = Namespace('health', description='Health check operations')
ns_config = Namespace('config', description='Configuration operations')
ns_quotes = Namespace('quotes', description='Insurance quotes operations')
ns_notif  = Namespace('notifications', description='Insurance claim email monitoring')

api.add_namespace(ns_health, path='/health')
api.add_namespace(ns_config, path='/config')
api.add_namespace(ns_quotes, path='/quotes')
api.add_namespace(ns_notif,  path='/notifications')

# Start Gmail IMAP inbox watcher
start_inbox_watcher()

# ============== SWAGGER MODELS ==============

# Health Response Model
health_model = api.model('HealthResponse', {
    'status': fields.String(description='API health status', example='healthy'),
    'timestamp': fields.String(description='Current timestamp', example='2026-02-17T10:30:00'),
    'version': fields.String(description='API version', example='1.0.0')
})

# Provider Model
provider_model = api.model('Provider', {
    'id': fields.String(description='Provider ID', example='icici'),
    'name': fields.String(description='Provider display name', example='ICICI Lombard'),
    'logo': fields.String(description='Provider logo URL', example='/logos/icici.png')
})

# Config Response Model
config_model = api.model('ConfigResponse', {
    'defaultVehicle': fields.String(description='Default vehicle registration number', example='KA02ML2085'),
    'defaultName': fields.String(description='Default customer name', example='Priyanka Shah'),
    'defaultMobile': fields.String(description='Default mobile number', example='6362493807'),
    'defaultEmail': fields.String(description='Default email address', example='priyankashah8324@gmail.com'),
    'providers': fields.List(fields.Nested(provider_model), description='Available insurance providers')
})

# Quote Request Model
quote_request_model = api.model('QuoteRequest', {
    'vehicleNumber': fields.String(required=True, description='Vehicle registration number', example='KA02ML2085'),
    'fullName': fields.String(required=True, description='Customer full name', example='Priyanka Shah'),
    'mobile': fields.String(required=True, description='10-digit mobile number', example='6362493807'),
    'email': fields.String(required=True, description='Email address', example='priyankashah8324@gmail.com'),
    'providers': fields.List(fields.String, required=True, description='List of providers to fetch quotes from',
                             example=['ICICI Lombard', 'PolicyBazaar', 'InsuranceDekho'])
})

# AI Recommendation Model
recommendation_model = api.model('AIRecommendation', {
    'rating': fields.String(description='Rating category', example='HIGHLY_RECOMMENDED'),
    'stars': fields.Integer(description='Star rating (1-5)', example=5),
    'highlights': fields.List(fields.String, description='Key highlights', 
                              example=['💰 Excellent value - competitive pricing', '🏥 Excellent network - 5000+ cashless garages']),
    'score': fields.Integer(description='Overall score', example=8)
})

# Single Quote Model
quote_model = api.model('Quote', {
    'provider': fields.String(description='Provider source', example='ICICI Lombard'),
    'insuranceProvider': fields.String(description='Insurance company name', example='ICICI Lombard General Insurance'),
    'planName': fields.String(description='Plan/package name', example='Comprehensive Plan'),
    'annualPremium': fields.Integer(description='Annual premium in INR', example=3500),
    'idv': fields.String(description='Insured Declared Value', example='₹1,50,000'),
    'cashlessGarages': fields.String(description='Number of cashless garages', example='5,200'),
    'keyBenefits': fields.List(fields.String, description='List of key benefits',
                               example=['Zero Depreciation', 'Roadside Assistance', '24x7 Claim Support']),
    'recommendation': fields.Nested(recommendation_model, description='AI-generated recommendation'),
    'status': fields.String(description='Quote status', enum=['success', 'partial', 'error'], example='success'),
    'error': fields.String(description='Error message if status is error', example=None)
})

# Provider Status Model
provider_status_model = api.model('ProviderStatus', {
    'provider': fields.String(description='Provider name', example='ICICI Lombard'),
    'status': fields.String(description='Fetch status', enum=['success', 'error'], example='success'),
    'quotesCount': fields.Integer(description='Number of quotes fetched', example=3),
    'error': fields.String(description='Error message if failed', example=None)
})

# Summary Model
summary_model = api.model('QuotesSummary', {
    'totalQuotes': fields.Integer(description='Total successful quotes', example=5),
    'lowestPremium': fields.Integer(description='Lowest premium found', example=3200),
    'highestPremium': fields.Integer(description='Highest premium found', example=6500),
    'providersChecked': fields.Integer(description='Number of providers checked', example=3),
    'successfulProviders': fields.Integer(description='Providers that returned quotes', example=2),
    'failedProviders': fields.Integer(description='Providers that failed', example=1)
})

# Quote Response Model
quote_response_model = api.model('QuoteResponse', {
    'success': fields.Boolean(description='Operation success status', example=True),
    'timestamp': fields.String(description='Response timestamp', example='2026-02-17T10:30:00'),
    'vehicleNumber': fields.String(description='Vehicle registration number', example='KA02ML2085'),
    'summary': fields.Nested(summary_model, description='Summary of quotes'),
    'providerStatus': fields.List(fields.Nested(provider_status_model), description='Status of each provider'),
    'quotes': fields.List(fields.Nested(quote_model), description='List of insurance quotes')
})

# Error Response Model
error_model = api.model('ErrorResponse', {
    'success': fields.Boolean(description='Always false for errors', example=False),
    'errors': fields.List(fields.String, description='List of error messages',
                          example=['Invalid vehicle registration number', 'Mobile number must be 10 digits'])
})

# Export Request Model
export_request_model = api.model('ExportRequest', {
    'quotes': fields.List(fields.Nested(quote_model), required=True, description='Quotes to export'),
    'vehicleNumber': fields.String(required=True, description='Vehicle registration number', example='KA02ML2085')
})

# Export Response Model
export_response_model = api.model('ExportResponse', {
    'success': fields.Boolean(description='Export success status', example=True),
    'filename': fields.String(description='Generated filename', example='insurance_quotes_20260217_103000.xlsx'),
    'downloadUrl': fields.String(description='URL to download file', example='/api/download/insurance_quotes_20260217_103000.xlsx')
})

# ============== NOTIFICATION MODELS ==============

claim_summary_model = ns_notif.model('ClaimSummary', {
    'total': fields.Integer(required=True, description='Total claims', example=10),
    'approved': fields.Integer(required=True, description='Approved claims', example=5),
    'rejected': fields.Integer(required=True, description='Rejected claims', example=2),
    'pending': fields.Integer(required=True, description='Pending claims', example=3),
})

claim_model = ns_notif.model('Claim', {
    'safe_id': fields.String(required=True, description='Unique and safe ID for the claim email'),
    'claim_id': fields.String(description='Claim reference ID'),
    'policy_number': fields.String(description='Policy number associated with the claim'),
    'status': fields.String(description='Current status of the claim (e.g., pending, approved, rejected)'),
    'subject': fields.String(description='Email subject line'),
    'from_addr': fields.String(description='Sender email address'),
    'submitted_by': fields.String(description='Name of the person who submitted the claim'),
    'date_submitted': fields.String(description='Date the claim was submitted'),
    'body_preview': fields.String(description='A short preview of the email body'),
})

pagination_model = ns_notif.model('Pagination', {
    'page': fields.Integer(required=True, description='Current page number'),
    'per_page': fields.Integer(required=True, description='Items per page'),
    'total_items': fields.Integer(required=True, description='Total number of items'),
    'total_pages': fields.Integer(required=True, description='Total number of pages'),
})

claims_dashboard_model = ns_notif.model('ClaimsDashboard', {
    'success': fields.Boolean(required=True, description='Indicates if the request was successful'),
    'summary': fields.Nested(claim_summary_model, required=True),
    'claims': fields.List(fields.Nested(claim_model), required=True),
    'pagination': fields.Nested(pagination_model, required=True),
})

watcher_status_model = ns_notif.model('WatcherStatus', {
    'success': fields.Boolean(required=True),
    'watcher_running': fields.Boolean(required=True),
    'monitored_inbox': fields.String(),
    'poll_interval_sec': fields.Integer(),
    'unhandled_count': fields.Integer(),
})

# Store job status and results
jobs = {}

def generate_ai_recommendation(premium, garages, idv, benefits):
    """Generate AI-powered recommendation based on quote details"""
    recommendations = []
    score = 0
    
    # Analyze premium
    try:
        premium_val = int(premium) if premium else 0
        if premium_val > 0:
            if premium_val < 3500:
                recommendations.append("💰 Excellent value - competitive pricing")
                score += 3
            elif premium_val < 4500:
                recommendations.append("💰 Good value - reasonable pricing")
                score += 2
            elif premium_val < 6000:
                recommendations.append("💰 Standard pricing")
                score += 1
            else:
                recommendations.append("💰 Premium pricing")
    except:
        pass
    
    # Analyze cashless garages
    try:
        garages_str = str(garages).replace(",", "").replace("N/A", "0")
        garages_val = int(garages_str) if garages_str.isdigit() else 0
        if garages_val >= 5000:
            recommendations.append("🏥 Excellent network - 5000+ cashless garages")
            score += 3
        elif garages_val >= 3000:
            recommendations.append("🏥 Good network - 3000+ cashless garages")
            score += 2
        elif garages_val >= 1000:
            recommendations.append("🏥 Moderate network coverage")
            score += 1
    except:
        pass
    
    # Analyze IDV
    try:
        idv_str = str(idv).replace(",", "").replace("N/A", "0").replace("₹", "")
        idv_val = int(idv_str) if idv_str.isdigit() else 0
        if idv_val >= 150000:
            recommendations.append("🛡️ High IDV coverage")
            score += 2
        elif idv_val >= 100000:
            recommendations.append("🛡️ Standard IDV coverage")
            score += 1
    except:
        pass
    
    # Analyze benefits
    if benefits:
        benefit_keywords = {
            "roadside": "🚗 Roadside assistance",
            "towing": "🚗 Towing service",
            "zero depreciation": "⭐ Zero depreciation",
            "cashless": "✅ Cashless claims",
            "pick up": "🚙 Free pick up & drop",
            "warranty": "🔧 Repair warranty"
        }
        for keyword, msg in benefit_keywords.items():
            for benefit in benefits:
                if keyword.lower() in str(benefit).lower():
                    if msg not in recommendations:
                        recommendations.append(msg)
                        score += 1
                    break
    
    # Generate rating
    if score >= 7:
        rating = "HIGHLY_RECOMMENDED"
        stars = 5
    elif score >= 5:
        rating = "RECOMMENDED"
        stars = 4
    elif score >= 3:
        rating = "GOOD_OPTION"
        stars = 3
    elif score >= 1:
        rating = "CONSIDER"
        stars = 2
    else:
        rating = "BASIC"
        stars = 1
    
    return {
        "rating": rating,
        "stars": stars,
        "highlights": recommendations,
        "score": score
    }


def parse_result(provider, result_text):
    """Parse the automation result into structured data"""
    try:
        # Clean up the result text
        if "```json" in result_text:
            result_text = result_text.split("```json")[1].split("```")[0].strip()
        elif "```" in result_text:
            result_text = result_text.split("```")[1].split("```")[0].strip()
        
        result_text = result_text.replace("\\n", "").replace('\\"', '"')
        if result_text.startswith('"""'):
            result_text = result_text[3:-3].strip()
        
        data = json.loads(result_text)
        quotes = data.get("quotes", [])
        
        parsed_quotes = []
        for quote in quotes:
            benefits = quote.get("key_benefits", [])
            recommendation = generate_ai_recommendation(
                quote.get("annual_premium_inr", 0),
                quote.get("cashless_garages", "N/A"),
                quote.get("idv", "N/A"),
                benefits
            )
            
            parsed_quotes.append({
                "provider": provider,
                "insuranceProvider": quote.get("insurance_provider", provider),
                "planName": quote.get("plan_name", "N/A"),
                "annualPremium": quote.get("annual_premium_inr", 0),
                "idv": quote.get("idv", "N/A"),
                "cashlessGarages": quote.get("cashless_garages", "N/A"),
                "keyBenefits": benefits,
                "recommendation": recommendation,
                "status": "success"
            })
        
        return parsed_quotes if parsed_quotes else [{
            "provider": provider,
            "insuranceProvider": provider,
            "planName": "No plans extracted",
            "annualPremium": 0,
            "idv": "N/A",
            "cashlessGarages": "N/A",
            "keyBenefits": [],
            "recommendation": {"rating": "N/A", "stars": 0, "highlights": [], "score": 0},
            "status": "partial",
            "rawData": result_text[:500]
        }]
        
    except Exception as e:
        return [{
            "provider": provider,
            "insuranceProvider": provider,
            "planName": "Parse Error",
            "annualPremium": 0,
            "idv": "N/A",
            "cashlessGarages": "N/A",
            "keyBenefits": [],
            "recommendation": {"rating": "N/A", "stars": 0, "highlights": [], "score": 0},
            "status": "error",
            "error": str(e)
        }]


async def run_provider(provider):
    """Run a single provider automation"""
    try:
        if provider == "ICICI Lombard":
            result = await run_icici_lombard()
        elif provider == "PolicyBazaar":
            result = await run_policybazaar()
        elif provider == "InsuranceDekho":
            result = await run_insurancedekho()
        else:
            return {"provider": provider, "status": "error", "error": "Unknown provider"}
        
        return {"provider": provider, "status": "success", "data": result}
    except Exception as e:
        return {"provider": provider, "status": "error", "error": str(e)}


def run_async(coro):
    """Run async function in sync context"""
    if sys.platform == 'win32':
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    try:
        return loop.run_until_complete(coro)
    finally:
        loop.close()


# ============== API ENDPOINTS ==============

# Root endpoint (outside namespace)
@app.route('/', methods=['GET'])
def index():
    """Root endpoint - API welcome page"""
    return {
        "name": "Insurance Quote API",
        "version": "1.0.0",
        "status": "running",
        "swagger_ui": "/swagger",
        "endpoints": {
            "health": "GET /api/health",
            "config": "GET /api/config",
            "quotes": "POST /api/quotes",
            "export_excel": "POST /api/quotes/export",
            "export_pdf": "POST /api/quotes/export/pdf",
            "download": "GET /api/download/<filename>"
        },
        "frontend": "https://www.insurebot.com",
        "message": "API is running! Visit /swagger for interactive documentation."
    }


@ns_health.route('')
class HealthCheck(Resource):
    @ns_health.doc('health_check')
    @ns_health.marshal_with(health_model)
    def get(self):
        """Check API health status"""
        return {
            "status": "healthy",
            "timestamp": datetime.now().isoformat(),
            "version": "1.0.0"
        }


@ns_config.route('')
class Config(Resource):
    @ns_config.doc('get_config')
    @ns_config.marshal_with(config_model)
    def get(self):
        """Get default configuration and available providers"""
        return {
            "defaultVehicle": VEHICLE_REG_NO,
            "defaultName": FULL_NAME,
            "defaultMobile": VALID_USERNAME,
            "defaultEmail": EMAIL_ID,
            "providers": [
                {"id": "icici", "name": "ICICI Lombard", "logo": "/logos/icici.png"},
                {"id": "policybazaar", "name": "PolicyBazaar", "logo": "/logos/policybazaar.png"},
                {"id": "insurancedekho", "name": "InsuranceDekho", "logo": "/logos/insurancedekho.png"}
            ]
        }


@ns_quotes.route('')
class Quotes(Resource):
    @ns_quotes.doc('get_quotes')
    @ns_quotes.expect(quote_request_model)
    @ns_quotes.response(200, 'Success', quote_response_model)
    @ns_quotes.response(400, 'Validation Error', error_model)
    def post(self):
        """
        Get insurance quotes from selected providers
        
        Submit vehicle and customer details to fetch quotes from multiple insurance providers.
        The API will run automation for each selected provider and return consolidated quotes.
        """
        data = request.json
        
        # Validation
        vehicle_no = data.get('vehicleNumber', '').strip()
        name = data.get('fullName', '').strip()
        mobile = data.get('mobile', '').strip()
        email = data.get('email', '').strip()
        providers = data.get('providers', [])
        
        errors = []
        if not vehicle_no or len(vehicle_no) < 6:
            errors.append("Invalid vehicle registration number")
        if not name or len(name) < 2:
            errors.append("Full name is required")
        if not mobile or len(mobile) != 10 or not mobile.isdigit():
            errors.append("Mobile number must be 10 digits")
        if not email or "@" not in email:
            errors.append("Valid email is required")
        if not providers:
            errors.append("Select at least one provider")
        
        if errors:
            return {"success": False, "errors": errors}, 400
        
        # Run automation for each provider
        all_quotes = []
        provider_status = []
        
        for provider in providers:
            result = run_async(run_provider(provider))
            
            if result["status"] == "success":
                quotes = parse_result(provider, result["data"])
                all_quotes.extend(quotes)
                provider_status.append({
                    "provider": provider,
                    "status": "success",
                    "quotesCount": len(quotes)
                })
            else:
                all_quotes.append({
                    "provider": provider,
                    "insuranceProvider": provider,
                    "planName": "Error",
                    "annualPremium": 0,
                    "idv": "N/A",
                    "cashlessGarages": "N/A",
                    "keyBenefits": [],
                    "recommendation": {"rating": "N/A", "stars": 0, "highlights": [], "score": 0},
                    "status": "error",
                    "error": result.get("error", "Unknown error")
                })
                provider_status.append({
                    "provider": provider,
                    "status": "error",
                    "error": result.get("error", "Unknown error")
                })
        
        # Sort by premium (lowest first)
        all_quotes.sort(key=lambda x: x.get("annualPremium", 0) or 999999)
        
        # Calculate summary
        successful_quotes = [q for q in all_quotes if q["status"] == "success"]
        premiums = [q["annualPremium"] for q in successful_quotes if q["annualPremium"] > 0]
        
        return {
            "success": True,
            "timestamp": datetime.now().isoformat(),
            "vehicleNumber": vehicle_no,
            "summary": {
                "totalQuotes": len(successful_quotes),
                "lowestPremium": min(premiums) if premiums else 0,
                "highestPremium": max(premiums) if premiums else 0,
                "providersChecked": len(providers),
                "successfulProviders": len([p for p in provider_status if p["status"] == "success"]),
                "failedProviders": len([p for p in provider_status if p["status"] == "error"])
            },
            "providerStatus": provider_status,
            "quotes": all_quotes
        }


@ns_quotes.route('/export')
class ExportExcel(Resource):
    @ns_quotes.doc('export_excel')
    @ns_quotes.expect(export_request_model)
    @ns_quotes.response(200, 'Success', export_response_model)
    @ns_quotes.response(400, 'Bad Request', error_model)
    def post(self):
        """
        Export quotes to Excel file
        
        Generate a formatted Excel file with all quote details including:
        - Provider information
        - Plan details and premiums
        - AI recommendations
        - Key benefits
        """
        import pandas as pd
        from openpyxl.styles import Font, PatternFill, Alignment
        
        data = request.json
        quotes = data.get('quotes', [])
        vehicle_no = data.get('vehicleNumber', 'unknown')
        
        if not quotes:
            return {"success": False, "errors": ["No quotes to export"]}, 400
        
        # Convert to Excel format
        excel_data = []
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        for quote in quotes:
            benefits = quote.get("keyBenefits", [])
            benefits_text = "\n• " + "\n• ".join(benefits) if benefits else ""
            
            rec = quote.get("recommendation", {})
            rec_text = f"{'⭐' * rec.get('stars', 0)} {rec.get('rating', 'N/A')}\n\n" + "\n".join(rec.get('highlights', []))
            
            excel_data.append({
                "Provider": quote.get("provider", ""),
                "Vehicle": vehicle_no,
                "Plan Name": quote.get("planName", ""),
                "Insurance Provider": quote.get("insuranceProvider", ""),
                "Annual Premium (INR)": quote.get("annualPremium", 0),
                "IDV": quote.get("idv", ""),
                "Cashless Garages": quote.get("cashlessGarages", ""),
                "Key Benefits": benefits_text,
                "AI Recommendation": rec_text,
                "Status": quote.get("status", ""),
                "Timestamp": timestamp
            })
        
        # Create DataFrame and save
        df = pd.DataFrame(excel_data)
        df = df.sort_values(by="Annual Premium (INR)")
        
        filename = f"insurance_quotes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        filepath = os.path.join(os.path.dirname(__file__), filename)
        
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Insurance Quotes', index=False)
            worksheet = writer.sheets['Insurance Quotes']
            
            # Format columns
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if cell.value:
                            cell_length = len(str(cell.value))
                            if max_length < cell_length:
                                max_length = cell_length
                    except:
                        pass
                adjusted_width = min(max_length + 2, 80)
                worksheet.column_dimensions[column_letter].width = adjusted_width
            
            # Header styling
            header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            
            for cell in worksheet[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
        
        return {
            "success": True,
            "filename": filename,
            "downloadUrl": f"/api/download/{filename}"
        }


@ns_quotes.route('/export/pdf')
class ExportPDF(Resource):
    @ns_quotes.doc('export_pdf')
    @ns_quotes.expect(export_request_model)
    @ns_quotes.response(200, 'Success', export_response_model)
    @ns_quotes.response(400, 'Bad Request', error_model)
    @ns_quotes.response(500, 'Server Error')
    def post(self):
        """
        Export quotes to PDF file
        
        Generate a professionally formatted PDF report with:
        - Quote comparison table
        - Summary statistics
        - AI recommendations
        - Visual styling and branding
        """
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import letter, A4
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import inch
            from reportlab.lib.enums import TA_CENTER, TA_LEFT
            
            data = request.json
            quotes = data.get('quotes', [])
            vehicle_no = data.get('vehicleNumber', 'unknown')
            
            if not quotes:
                return {"success": False, "errors": ["No quotes to export"]}, 400
            
            # Create PDF filename
            filename = f"insurance_quotes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            filepath = os.path.join(os.path.dirname(__file__), filename)
            
            # Create PDF document
            doc = SimpleDocTemplate(filepath, pagesize=A4, 
                                  rightMargin=30, leftMargin=30,
                                  topMargin=30, bottomMargin=18)
            
            # Container for elements
            elements = []
            styles = getSampleStyleSheet()
            
            # Custom styles
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=24,
                textColor=colors.HexColor('#1e40af'),
                spaceAfter=30,
                alignment=TA_CENTER,
                fontName='Helvetica-Bold'
            )
            
            subtitle_style = ParagraphStyle(
                'CustomSubtitle',
                parent=styles['Normal'],
                fontSize=12,
                textColor=colors.HexColor('#6b7280'),
                spaceAfter=20,
                alignment=TA_CENTER
            )
            
            # Add title
            title = Paragraph("Insurance Quote Comparison Report", title_style)
            elements.append(title)
            
            # Add subtitle with vehicle number and timestamp
            timestamp = datetime.now().strftime("%B %d, %Y at %I:%M %p")
            subtitle = Paragraph(f"Vehicle: <b>{vehicle_no}</b> | Generated: {timestamp}", subtitle_style)
            elements.append(subtitle)
            elements.append(Spacer(1, 20))
            
            # Sort quotes by premium
            sorted_quotes = sorted(quotes, key=lambda x: x.get('annualPremium', 0) if x.get('annualPremium') else float('inf'))
            
            # Prepare table data
            table_data = [['Provider', 'Plan', 'Premium\n(INR/year)', 'IDV', 'Garages', 'Rating', 'Status']]
            
            for quote in sorted_quotes:
                rec = quote.get('recommendation', {})
                stars = '*' * rec.get('stars', 0) if rec.get('stars') else 'N/A'
                rating = rec.get('rating', 'N/A').replace('_', ' ')
                
                status_text = {
                    'success': 'OK',
                    'partial': 'PARTIAL',
                    'error': 'ERROR'
                }.get(quote.get('status', ''), '')
                
                row = [
                    quote.get('provider', 'N/A'),
                    quote.get('planName', 'N/A')[:30],
                    f"Rs.{quote.get('annualPremium', 0):,}" if quote.get('annualPremium') else 'N/A',
                    str(quote.get('idv', 'N/A')),
                    str(quote.get('cashlessGarages', 'N/A')),
                    f"{stars}\n{rating}"[:40],
                    f"{status_text}"
                ]
                table_data.append(row)
            
            # Create table
            col_widths = [1.2*inch, 1.8*inch, 1.1*inch, 0.9*inch, 0.8*inch, 1.3*inch, 0.9*inch]
            table = Table(table_data, colWidths=col_widths)
            
            # Table styling
            table.setStyle(TableStyle([
                # Header
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1e40af')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                
                # Data rows
                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                ('ALIGN', (0, 1), (-1, -1), 'LEFT'),
                ('ALIGN', (2, 1), (2, -1), 'RIGHT'),  # Premium column right-aligned
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('TOPPADDING', (0, 1), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                
                # Alternating row colors
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f3f4f6')]),
            ]))
            
            elements.append(table)
            elements.append(Spacer(1, 30))
            
            # Add summary
            successful_quotes = [q for q in sorted_quotes if q.get('status') == 'success' and q.get('annualPremium')]
            if successful_quotes:
                premiums = [q['annualPremium'] for q in successful_quotes]
                min_premium = min(premiums)
                max_premium = max(premiums)
                avg_premium = sum(premiums) // len(premiums)
                
                summary_style = ParagraphStyle(
                    'Summary',
                    parent=styles['Normal'],
                    fontSize=11,
                    textColor=colors.HexColor('#374151'),
                    spaceAfter=6,
                    leftIndent=20
                )
                
                summary_title = Paragraph("<b>Summary:</b>", summary_style)
                elements.append(summary_title)
                
                summary_points = [
                    f"<b>Total Plans Found:</b> {len(successful_quotes)}",
                    f"<b>Lowest Premium:</b> Rs.{min_premium:,}/year",
                    f"<b>Highest Premium:</b> Rs.{max_premium:,}/year",
                    f"<b>Average Premium:</b> Rs.{avg_premium:,}/year",
                ]
                
                for point in summary_points:
                    p = Paragraph(f"* {point}", summary_style)
                    elements.append(p)
            
            elements.append(Spacer(1, 20))
            
            # Footer note
            footer_style = ParagraphStyle(
                'Footer',
                parent=styles['Normal'],
                fontSize=9,
                textColor=colors.HexColor('#9ca3af'),
                alignment=TA_CENTER
            )
            footer = Paragraph("Generated by InsureBot - Your data is safe with us", footer_style)
            elements.append(footer)
            
            # Build PDF
            doc.build(elements)
            
            return {
                "success": True,
                "filename": filename,
                "downloadUrl": f"/api/download/{filename}"
            }
            
        except ImportError:
            return {
                "success": False, 
                "errors": ["PDF generation library not installed. Please install reportlab: pip install reportlab"]
            }, 500
        except Exception as e:
            return {"success": False, "errors": [str(e)]}, 500


# Download endpoint (outside namespace for cleaner URL)
@app.route('/api/download/<filename>', methods=['GET'])
def download_file(filename):
    """
    Download generated file (Excel or PDF)
    ---
    parameters:
      - name: filename
        in: path
        type: string
        required: true
        description: Name of the file to download
    responses:
      200:
        description: File download
      404:
        description: File not found
    """
    filepath = os.path.join(os.path.dirname(__file__), filename)
    
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    else:
        return {"success": False, "error": "File not found"}, 404


# ── /api/notifications endpoints (Gmail inbox missed-email monitor) ───────────

claim_body_model = ns_notif.model('ClaimEmailBody', {
    'success':    fields.Boolean(),
    'claim_id':   fields.String(description='Claim reference extracted from the email'),
    'from_addr':  fields.String(description='Sender email address'),
    'subject':    fields.String(description='Email subject line'),
    'body':       fields.String(description='Full plain-text body of the email'),
    'claim_info': fields.Raw(description='All extracted claim fields (status, amounts, dates, …)'),
})

inbox_status_model = ns_notif.model('InboxStatus', {
    'success':           fields.Boolean(),
    'monitored_inbox':   fields.String(),
    'poll_interval_sec': fields.Integer(),
    'unhandled_count':   fields.Integer(description='Claim emails not yet handled'),
    'watcher_running':   fields.Boolean(),
})

# ── Claims Dashboard Swagger models ──────────────────────────────────────────────
dashboard_summary_model = ns_notif.model('DashboardSummary', {
    'total':    fields.Integer(description='Total claim emails (all statuses)', example=72),
    'approved': fields.Integer(description='Approved claims count', example=32),
    'rejected': fields.Integer(description='Rejected claims count', example=25),
    'pending':  fields.Integer(description='Pending / under-review claims count', example=15),
})

dashboard_pagination_model = ns_notif.model('DashboardPagination', {
    'page':        fields.Integer(description='Current page (1-based)', example=1),
    'per_page':    fields.Integer(description='Items per page', example=10),
    'total_items': fields.Integer(description='Total items matching filters', example=50),
    'total_pages': fields.Integer(description='Total pages', example=5),
})

dashboard_claim_model = ns_notif.model('DashboardClaim', {
    'claim_id':       fields.String(description='Extracted claim reference or short safe_id', example='CLM-78451239'),
    'safe_id':        fields.String(description='Internal email safe_id — pass to /claims-dashboard/{safe_id}/body to read full email'),
    'submitted_by':   fields.String(description='Customer name or sender address', example='Sarah Miller'),
    'status':         fields.String(description='Dashboard status bucket', enum=['approved', 'rejected', 'pending'], example='rejected'),
    'status_label':   fields.String(description='Human-readable status', example='Rejected'),
    'reason':         fields.String(description='Reason text shown in the dashboard', example='Denied due to policy limits'),
    'date_submitted': fields.String(description='Date of intimation or email received date', example='29 Nov 2021  14:05 UTC'),
    'handled':        fields.Boolean(description='Whether acknowledged by the manager', example=False),
    'policy_number':  fields.String(description='Extracted policy number (if any)', example='MOT-IND-45892173'),
    'from_addr':      fields.String(description='Raw sender email address', example='claims@insurer.com'),
    'subject':        fields.String(description='Email subject line'),
    'body_preview':   fields.String(description='First 300 characters of the email body — for inline preview in the dashboard'),
})

dashboard_response_model = ns_notif.model('ClaimsDashboardResponse', {
    'success':    fields.Boolean(example=True),
    'summary':    fields.Nested(dashboard_summary_model, description='Aggregate counts — always reflects ALL claims'),
    'claims':     fields.List(fields.Nested(dashboard_claim_model), description='Paginated, filtered claim rows'),
    'pagination': fields.Nested(dashboard_pagination_model),
})


@ns_notif.route('/claims-dashboard')
class ClaimsDashboard(Resource):
    @ns_notif.doc(
        'get_claims_dashboard',
        params={
            'status':   {'description': 'Filter by status bucket: all | approved | rejected | pending',
                         'in': 'query', 'type': 'string', 'default': 'all'},
            'search':   {'description': 'Free-text search across Claim ID, Submitted By, Reason, Subject',
                         'in': 'query', 'type': 'string', 'default': ''},
            'page':     {'description': '1-based page number', 'in': 'query', 'type': 'integer', 'default': 1},
            'per_page': {'description': 'Items per page (max 100)', 'in': 'query', 'type': 'integer', 'default': 10},
            'all':      {'description': 'Set to 1 to include already-handled emails in summary & list',
                         'in': 'query', 'type': 'integer', 'default': 0},
        }
    )
    @ns_notif.response(200, 'Success', dashboard_response_model)
    def get(self):
        """
        **Claims Review Dashboard** — summary cards + paginated claims table.

        Returns three sections that map directly to the frontend dashboard:

        | Section       | Purpose |
        |---------------|---------|
        | `summary`     | Aggregate counts for the three status cards (Approved / Rejected / Pending). Always computed over *all* stored claims — not affected by page/search filters. |
        | `claims`      | Paginated list of claim rows ready for the table. |
        | `pagination`  | `page`, `per_page`, `total_items`, `total_pages` |

        **Query parameters**

        * `status`   — `all` *(default)*, `approved`, `rejected`, `pending`
        * `search`   — substring match on Claim ID, Submitted By, Reason, or Subject
        * `page`     — 1-based page (default `1`)
        * `per_page` — rows per page (default `10`, max `100`)
        * `all=1`    — include already-handled / acknowledged emails (default: unhandled only)

        **Status mapping**

        | Raw claim_status (email) | Dashboard bucket |
        |--------------------------|-----------------|
        | `approved`               | approved        |
        | `partial`                | approved        |
        | `rejected`               | rejected        |
        | `under_review`           | pending         |
        | `intimated`              | pending         |
        | `unknown`                | pending         |
        """
        status_filter   = request.args.get('status',   'all').strip().lower()
        search          = request.args.get('search',   '').strip()
        include_handled = request.args.get('all', '0') == '1'

        try:
            page     = max(1, int(request.args.get('page',     1)))
            per_page = max(1, min(100, int(request.args.get('per_page', 10))))
        except (ValueError, TypeError):
            page, per_page = 1, 10

        data = get_claims_dashboard(
            status_filter   = status_filter,
            search          = search,
            page            = page,
            per_page        = per_page,
            include_handled = include_handled,
        )

        return {'success': True, **data}, 200




@ns_notif.route('/claims-dashboard/<string:safe_id>/body')
class ClaimEmailBody(Resource):
    @ns_notif.doc('get_claim_email_body')
    @ns_notif.response(200, 'Success', claim_body_model)
    @ns_notif.response(404, 'Claim not found')
    def get(self, safe_id):
        """
        Read the **full email body** for a specific claim row in the dashboard.

        Use the `safe_id` returned by each row in the `/claims-dashboard` response.
        Returns the complete plain-text message plus all extracted claim fields,
        so the frontend can show a rich detail panel / modal without a separate store.
        """
        entry = get_missed_email_by_id(safe_id)
        if entry is None:
            return {'success': False, 'error': f'Claim {safe_id!r} not found'}, 404
        ci       = entry.get('claim_info') or {}
        claim_id = ci.get('claim_ref') or safe_id[:10].upper()
        return {
            'success':    True,
            'claim_id':   claim_id,
            'from_addr':  entry.get('from_addr'),
            'subject':    entry.get('subject'),
            'body':       entry.get('full_body', ''),
            'claim_info': ci,
        }, 200


@ns_notif.route('/status')
class NotificationStatus(Resource):
    @ns_notif.doc('inbox_status')
    @ns_notif.response(200, 'Status', inbox_status_model)
    def get(self):
        """
        Health status of the Gmail inbox watcher.
        Returns how many unhandled claims are waiting and whether the
        background watcher thread is running.
        """
        dash = get_claims_dashboard(include_handled=False, per_page=1)
        return {
            'success':           True,
            'monitored_inbox':   GMAIL_USER,
            'poll_interval_sec': INBOX_POLL_SEC,
            'unhandled_count':   dash['summary']['total'],
            'watcher_running':   _inbox_watcher_started,
        }, 200


# ─────────────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    print("\n" + "=" * 60)
    print("Insurance Quote API Server")
    print("=" * 60)
    print("Server running at: https://api.insurebot.com (Production)")
    print("Local development: http://localhost:5000")
    print("Swagger UI: http://localhost:5000/swagger")
    print("=" * 60 + "\n")
    
    # Disable debug mode in production to prevent unnecessary restarts
    # Set use_reloader=False to prevent watchdog from monitoring site-packages
    app.run(host='127.0.0.1', port=5000, debug=False, use_reloader=False)

