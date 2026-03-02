"""
extractor_api.py

HTTP API wrapping doc_insights/extractor.py.
All business logic lives in extractor.py; this file handles routing only.

Endpoints:
  POST   /extractor/bulk-upload         Upload PDFs, extract & store in MongoDB
  GET    /extractor/policies             All stored policies sorted by expiry
  GET    /extractor/expiry-alerts        Policies expiring within ?days=90
  GET    /extractor/policy/<id>          Full details + compliance for one policy
  DELETE /extractor/policy/<id>          Remove policy and PDF from disk
  POST   /extractor/compliance-check     Check any PDF against compliance_report.pdf
  POST   /extractor/renewal-insights     Compare policy vs market quotes
"""

import os
import sys
import json
from datetime import datetime, timezone
from flask import Flask, request, send_file
from flask_cors import CORS
from flask_restx import Api, Resource, fields, reqparse
import werkzeug.datastructures
# Make doc_insights importable
sys.path.insert(0, os.path.dirname(__file__))

from doc_insights.extractor import (
    extract_text_from_pdf,
    extract_insights_with_llm,
    save_policy_to_db,
    process_bulk_pdfs,
    get_expiry_alerts,
    get_policy_insights,
    generate_renewal_insights,
    check_compliance_with_llm,
    COMPLIANCE_REPORT_PATH,
    _parse_expiry_date,
    _days_to_expiry,
    _alert_level,
    MONGO_URI,
    MONGO_DB,
    get_db,
    FALLBACK_JSON,
)
from bson import ObjectId
from bson.errors import InvalidId
from pymongo import ASCENDING


# Flask app
app = Flask(__name__)
CORS(app, origins="*", supports_credentials=False)

api = Api(
    app,
    version="1.0.0",
    title="Extractor API",
    description=(
        "Bulk-upload insurance PDFs → MongoDB Atlas → live expiry alerts & per-policy insights. "
        "All heavy logic lives in doc_insights/extractor.py."
    ),
    doc="/swagger",
)

ns = api.namespace("extractor", description="Insurance document operations")

# Permanent PDF storage folder
UPLOADS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
os.makedirs(UPLOADS_DIR, exist_ok=True)


# Swagger models
policy_summary_model = api.model("PolicySummary", {
    "id":                fields.String(description="MongoDB _id"),
    "policy_holder":     fields.String(),
    "vehicle":           fields.String(),
    "policy_type":       fields.String(),
    "policy_valid_till": fields.String(),
    "expiry_date":       fields.String(description="ISO datetime of expiry"),
    "days_to_expiry":    fields.Integer(description="Negative = already expired"),
    "alert_level":       fields.String(description="expired / critical / warning / ok"),
    "total_premium":     fields.String(),
    "idv":               fields.String(),
    "uploaded_at":       fields.String(),
    "filename":          fields.String(),
    "pdf_url":           fields.String(description="URL to view the original PDF in browser"),
})

policy_detail_model = api.model("PolicyDetail", {
    "id":                fields.String(),
    "policy_holder":     fields.String(),
    "vehicle":           fields.String(),
    "policy_type":       fields.String(),
    "policy_valid_till": fields.String(),
    "expiry_date":       fields.String(),
    "days_to_expiry":    fields.Integer(),
    "alert_level":       fields.String(),
    "total_premium":     fields.String(),
    "idv":               fields.String(),
    "ncb":               fields.String(),
    "add_ons":           fields.List(fields.String()),
    "coverage_details":  fields.List(fields.String()),
    "insights":          fields.List(fields.String(), description="Actionable insights shown on click"),
    "compliance":        fields.Raw(description="Compliance check result against compliance_report.pdf"),
    "uploaded_at":       fields.String(),
    "filename":          fields.String(),
    "pdf_url":           fields.String(description="URL to view the original PDF in browser"),
})

expiry_alerts_model = api.model("ExpiryAlertsResponse", {
    "success":        fields.Boolean(),
    "source":         fields.String(description="mongodb or local_json"),
    "threshold_days": fields.Integer(),
    "total_alerts":   fields.Integer(),
    "expired":        fields.Integer(),
    "critical":       fields.Integer(description="<= 30 days"),
    "warning":        fields.Integer(description="31-60 days"),
    "ok":             fields.Integer(description="> 60 days"),
    "policies":       fields.List(fields.Nested(policy_summary_model)),
})

bulk_result_model = api.model("BulkUploadResponse", {
    "success":  fields.Boolean(),
    "uploaded": fields.Integer(),
    "failed":   fields.Integer(),
    "results":  fields.List(fields.Raw()),
})

renewal_request_model = api.model("RenewalInsightsRequest", {
    "current_policy": fields.Raw(required=True,  description="Extracted policy dict"),
    "market_quotes":  fields.List(fields.Raw(), required=True, description="List of market quote dicts"),
    "user_city":      fields.String(required=False, default="Bengaluru"),
})

renewal_response_model = api.model("RenewalInsightsResponse", {
    "insights":   fields.List(fields.String()),
    "comparison": fields.List(fields.List(fields.String())),
})

compliance_result_model = api.model("ComplianceResult", {
    "compliant": fields.Boolean(description="True if no critical issues found"),
    "score":     fields.Integer(description="Compliance score 0-100"),
    "passed":    fields.List(fields.String(), description="Requirements that are satisfied"),
    "issues":    fields.List(fields.String(), description="Requirements that are missing or violated"),
    "summary":   fields.String(description="One-line overall verdict"),
})

compliance_response_model = api.model("ComplianceCheckResponse", {
    "success":    fields.Boolean(),
    "filename":   fields.String(),
    "compliance": fields.Nested(compliance_result_model),
})


# -- Helper: convert alert dict (from get_expiry_alerts) to summary dict ------
def _alert_to_summary(a):
    return {
        "id":                a["id"],
        "policy_holder":     a["policy_holder"],
        "vehicle":           a["vehicle"],
        "policy_type":       a["policy_type"],
        "policy_valid_till": a["policy_valid_till"],
        "expiry_date":       None,
        "days_to_expiry":    a["days_to_expiry"],
        "alert_level":       a["alert_level"],
        "total_premium":     a["total_premium"],
        "idv":               a["idv"],
        "uploaded_at":       a.get("_raw", {}).get("uploaded_at", "-"),
        "filename":          a["filename"],
    }


# Upload parsers (reqparse + FileStorage = Swagger shows a file picker)
_bulk_upload_parser = reqparse.RequestParser()
_bulk_upload_parser.add_argument(
    "files",
    location="files",
    type=werkzeug.datastructures.FileStorage,
    required=True,
    action="append",
    help="One or more insurance PDF files",
)

_compliance_upload_parser = reqparse.RequestParser()
_compliance_upload_parser.add_argument(
    "file",
    location="files",
    type=werkzeug.datastructures.FileStorage,
    required=True,
    help="Insurance PDF to check against compliance_report.pdf",
)


# Endpoints

@ns.route("/bulk-upload")
class BulkUpload(Resource):
    @ns.expect(_bulk_upload_parser)
    @ns.response(200, "Success", bulk_result_model)
    @ns.response(400, "Bad Request")
    def post(self):
        """
        Upload one or more insurance PDF files.
        Each PDF is extracted via LLM and stored in MongoDB Atlas.
        Re-uploading the same policy (vehicle + expiry) updates the record instead of duplicating.
        """

        # Accept both single and multiple file uploads
        import logging
        files = request.files.getlist("files")
        logging.warning(f"[UPLOAD DEBUG] files from getlist: {[f.filename for f in files]}")
        if not files or (len(files) == 1 and files[0].filename == ""):
            # Try single file (not as list)
            single_file = request.files.get("files")
            logging.warning(f"[UPLOAD DEBUG] single_file: {single_file.filename if single_file else None}")
            if single_file and single_file.filename:
                files = [single_file]
            else:
                logging.warning("[UPLOAD DEBUG] No files uploaded — send as multipart/form-data with field name 'files'")
                return {"success": False, "errors": ["No files uploaded — send as multipart/form-data with field name 'files'"]}, 400

        import hashlib
        pdf_paths = []
        saved_files = []
        duplicate_files = []
        existing_hashes = set()
        # Build hash set of already-uploaded files (by content)
        for existing in os.listdir(UPLOADS_DIR):
            existing_path = os.path.join(UPLOADS_DIR, existing)
            if os.path.isfile(existing_path) and existing.lower().endswith('.pdf'):
                try:
                    with open(existing_path, 'rb') as ef:
                        file_hash = hashlib.md5(ef.read()).hexdigest()
                        existing_hashes.add(file_hash)
                except Exception:
                    pass

        for f in files:
            fname = f.filename or "unknown.pdf"
            if not fname.lower().endswith(".pdf"):
                continue
            dest = os.path.join(UPLOADS_DIR, fname)
            # Check for duplicate by content hash
            f.seek(0)
            file_bytes = f.read()
            file_hash = hashlib.md5(file_bytes).hexdigest()
            if file_hash in existing_hashes:
                duplicate_files.append({"filename": fname, "reason": "Duplicate file content"})
                continue
            # Save file
            with open(dest, 'wb') as out_f:
                out_f.write(file_bytes)
            pdf_paths.append(dest)
            saved_files.append((dest, fname))
            existing_hashes.add(file_hash)

        if not pdf_paths:
            return {"success": False, "errors": ["No valid PDF files found"], "duplicates": duplicate_files}, 400

        # Run the tested bulk processor (PDFs remain on disk, file_path stored in MongoDB)
        results = process_bulk_pdfs(pdf_paths)

        # Set pdf_url in each result for the API response
        for i, (dest, orig_name) in enumerate(saved_files):
            if i < len(results):
                results[i]["filename"] = orig_name
                if results[i].get("status") == "success" and results[i].get("id"):
                    results[i]["pdf_url"] = f"/extractor/policy/{results[i]['id']}/pdf"

        uploaded = sum(1 for r in results if r.get("status") == "success")
        failed   = len(results) - uploaded

        # Ensure compliance dict is always present in each result
        for r in results:
            if r.get("status") == "success" and "compliance" not in r:
                r["compliance"] = {"compliant": None, "score": None, "passed": [], "issues": [], "summary": "-"}

        return {
            "success":  uploaded > 0,
            "uploaded": uploaded,
            "failed":   failed,
            "results":  results,
            "duplicates": duplicate_files,
        }, 200


@ns.route("/policies")
class PoliciesList(Resource):
    @ns.doc("list_policies")
    @ns.response(200, "Success")
    def get(self):
        """
        Return all stored policies sorted by expiry date (soonest first).
        Works from MongoDB Atlas; falls back to local JSON if MongoDB is down.
        """
        try:
            db   = get_db()
            docs = list(db.policies.find({}).sort("expiry_date", ASCENDING))
            policies = []
            for doc in docs:
                expiry_dt = doc.get("expiry_date")

                if isinstance(expiry_dt, datetime) and expiry_dt.tzinfo is None:
                    expiry_dt = expiry_dt.replace(tzinfo=timezone.utc)
                days  = _days_to_expiry(expiry_dt) if expiry_dt else _days_to_expiry(_parse_expiry_date(doc.get("policy_valid_till")))
                level = _alert_level(days)
                policies.append({
                    "id":                str(doc["_id"]),
                    "policy_holder":     doc.get("policy_holder", "-"),
                    "vehicle":           doc.get("vehicle", "-"),
                    "policy_type":       doc.get("policy_type", "-"),
                    "policy_valid_till": doc.get("policy_valid_till", "-"),
                    "expiry_date":       expiry_dt.isoformat() if expiry_dt else None,
                    "days_to_expiry":    days,
                    "alert_level":       level,
                    "total_premium":     doc.get("total_premium", "-"),
                    "idv":               doc.get("idv", "-"),
                    "uploaded_at":       doc.get("uploaded_at", "-"),
                    "filename":          doc.get("filename", "-"),
                    "pdf_url":           f"/extractor/policy/{str(doc['_id'])}/pdf" if doc.get("file_path") else None,
                })
            source = "mongodb"
        except Exception:
            if os.path.exists(FALLBACK_JSON):
                with open(FALLBACK_JSON, "r", encoding="utf-8") as f:
                    records = json.load(f)
            else:
                records = []
            policies = []
            for r in records:
                days  = _days_to_expiry(_parse_expiry_date(r.get("policy_valid_till")))
                level = _alert_level(days)
                policies.append({
                    "id":                r.get("_id", "-"),
                    "policy_holder":     r.get("policy_holder", "-"),
                    "vehicle":           r.get("vehicle", "-"),
                    "policy_type":       r.get("policy_type", "-"),
                    "policy_valid_till": r.get("policy_valid_till", "-"),
                    "expiry_date":       None,
                    "days_to_expiry":    days,
                    "alert_level":       level,
                    "total_premium":     r.get("total_premium", "-"),
                    "idv":               r.get("idv", "-"),
                    "uploaded_at":       r.get("uploaded_at", "-"),
                    "filename":          r.get("filename", "-"),
                    "pdf_url":           None,
                })
            policies.sort(key=lambda p: p.get("days_to_expiry") or 9999)
            source = "local_json"

        return {"success": True, "source": source, "total": len(policies), "policies": policies}, 200


@ns.route("/expiry-alerts")
class ExpiryAlerts(Resource):
    @ns.doc("expiry_alerts")
    @ns.response(200, "Success", expiry_alerts_model)
    def get(self):
        """
        Returns all policies expiring within the next N days (default 90) AND already-expired ones.
        days_to_expiry is recalculated live every request — always up to date.
        Query param: ?days=90
        """
        days_threshold = int(request.args.get("days", 90))

        # Calls the already-tested get_expiry_alerts() from extractor.py
        alerts = get_expiry_alerts(days_threshold)

        policies = [_alert_to_summary(a) for a in alerts]

        expired  = sum(1 for p in policies if p["alert_level"] == "expired")
        critical = sum(1 for p in policies if p["alert_level"] == "critical")
        warning  = sum(1 for p in policies if p["alert_level"] == "warning")
        ok       = sum(1 for p in policies if p["alert_level"] == "ok")

        # Detect source from first alert _raw (has "_id" string = local_json)
        source = "local_json"
        if alerts:
            raw_id = str(alerts[0].get("id", ""))
            source = "mongodb" if len(raw_id) == 24 and all(c in "0123456789abcdef" for c in raw_id) else "local_json"

        return {
            "success":        True,
            "source":         source,
            "threshold_days": days_threshold,
            "total_alerts":   len(policies),
            "expired":        expired,
            "critical":       critical,
            "warning":        warning,
            "ok":             ok,
            "policies":       policies,
        }, 200


@ns.route("/policy/<string:policy_id>")
class PolicyDetail(Resource):
    @ns.doc("policy_detail")
    @ns.response(200, "Success", policy_detail_model)
    @ns.response(404, "Not found")
    def get(self, policy_id):
        """
        Get full details + AI insights for ONE policy.
        This is called when the user clicks on an expiry alert card.
        Insights are generated fresh on every request by get_policy_insights().
        """
        try:
            oid = ObjectId(policy_id)
        except InvalidId:
            return {"success": False, "error": "Invalid policy ID format"}, 400

        try:
            db  = get_db()
            doc = db.policies.find_one({"_id": oid})
        except Exception as exc:
            return {"success": False, "error": f"Database error: {exc}"}, 503

        if not doc:
            return {"success": False, "error": f"Policy {policy_id!r} not found in MongoDB"}, 404

        # Build response
        expiry_dt = doc.get("expiry_date")
        if isinstance(expiry_dt, str):
            expiry_dt = _parse_expiry_date(expiry_dt)
        elif isinstance(expiry_dt, datetime) and expiry_dt.tzinfo is None:
            expiry_dt = expiry_dt.replace(tzinfo=timezone.utc)
        if expiry_dt is None:
            expiry_dt = _parse_expiry_date(doc.get("policy_valid_till"))

        days  = _days_to_expiry(expiry_dt)
        level = _alert_level(days)

        # get_policy_insights() is the tested function from extractor.py
        insights = get_policy_insights(doc)

        # Run compliance check — use ONLY the stored uploaded PDF file
        compliance = doc.get("compliance")  # use previously stored result if present
        if not compliance:
            file_path = doc.get("file_path")
            if file_path and os.path.exists(file_path):
                doc_text = extract_text_from_pdf(file_path)
                compliance = check_compliance_with_llm(doc_text, policy_dict=doc)
            else:
                compliance = {
                    "compliant": None,
                    "score": None,
                    "passed": [],
                    "issues": ["Original PDF not available on server. Re-upload the document to run compliance check."],
                    "summary": "Compliance check unavailable — PDF not found.",
                }

        return {
            "success": True,
            "policy": {
                "id":                str(doc.get("_id", policy_id)),
                "policy_holder":     doc.get("policy_holder", "-"),
                "vehicle":           doc.get("vehicle", "-"),
                "policy_type":       doc.get("policy_type", "-"),
                "policy_valid_till": doc.get("policy_valid_till", "-"),
                "expiry_date":       expiry_dt.isoformat() if expiry_dt else None,
                "days_to_expiry":    days,
                "alert_level":       level,
                "total_premium":     doc.get("total_premium", "-"),
                "idv":               doc.get("idv", "-"),
                "ncb":               doc.get("ncb", "-"),
                "add_ons":           doc.get("add_ons") or [],
                "coverage_details":  doc.get("coverage_details") or [],
                "insights":          insights,
                "compliance":        compliance,
                "uploaded_at":       doc.get("uploaded_at", "-"),
                "filename":          doc.get("filename", "-"),
                "pdf_url":           f"/extractor/policy/{str(doc.get('_id', policy_id))}/pdf" if doc.get("file_path") else None,
            }
        }, 200

    @ns.doc("delete_policy")
    @ns.response(200, "Deleted")
    @ns.response(404, "Not found")
    def delete(self, policy_id):
        """Delete a policy record from MongoDB and remove the PDF file from disk."""
        try:
            oid = ObjectId(policy_id)
        except InvalidId:
            return {"success": False, "error": "Invalid policy ID format"}, 400
        try:
            db  = get_db()
            doc = db.policies.find_one({"_id": oid}, {"file_path": 1})
            if not doc:
                return {"success": False, "error": f"Policy {policy_id!r} not found"}, 404
            # Remove physical PDF from disk
            file_path = doc.get("file_path")
            if file_path and os.path.exists(file_path):
                try:
                    os.remove(file_path)
                except Exception:
                    pass  # Don't fail the delete if file removal fails
            result = db.policies.delete_one({"_id": oid})
        except Exception as exc:
            return {"success": False, "error": f"Database error: {exc}"}, 503
        if result.deleted_count == 0:
            return {"success": False, "error": f"Policy {policy_id!r} not found"}, 404
        return {"success": True, "message": "Policy and PDF file deleted"}, 200


@ns.route("/policy/<string:policy_id>/pdf")
class PolicyPDF(Resource):
    @ns.doc("policy_pdf")
    @ns.response(200, "PDF file")
    @ns.response(404, "Not found")
    def get(self, policy_id):
        """
        Stream the original uploaded PDF for a policy.
        Opens inline in the browser — no download required.
        """
        try:
            oid = ObjectId(policy_id)
            db  = get_db()
            doc = db.policies.find_one({"_id": oid}, {"file_path": 1, "filename": 1})
            if not doc:
                return {"success": False, "error": "Policy not found"}, 404

            file_path = doc.get("file_path")
            if not file_path or not os.path.exists(file_path):
                return {"success": False, "error": "PDF not available on server. Re-upload the document to enable PDF viewing."}, 404

            return send_file(
                file_path,
                mimetype="application/pdf",
                as_attachment=False,
                download_name=doc.get("filename", "policy.pdf"),
            )
        except InvalidId:
            return {"success": False, "error": "Invalid policy ID"}, 400
        except Exception as e:
            return {"success": False, "error": str(e)}, 500


@ns.route("/compliance-check")
class ComplianceCheck(Resource):
    @ns.expect(_compliance_upload_parser)
    @ns.response(200, "Success", compliance_response_model)
    @ns.response(400, "Bad Request")
    def post(self):
        """
        Upload any insurance PDF and check it against the compliance standards
        defined in doc_insights/compliance_report.pdf.

        Returns a compliance result with score (0-100), list of passed requirements,
        list of issues / violations, and an overall verdict.
        """
        file = request.files.get("file")
        if not file:
            return {"success": False, "error": "No file uploaded — send as multipart/form-data with field name 'file'"}, 400

        fname = file.filename or "uploaded.pdf"
        if not fname.lower().endswith(".pdf"):
            return {"success": False, "error": "Only PDF files are supported"}, 400

        # Save to UPLOADS_DIR for text extraction
        dest = os.path.join(UPLOADS_DIR, fname)
        file.save(dest)

        # Extract text ONLY from the uploaded document
        doc_text = extract_text_from_pdf(dest)
        if not doc_text.strip():
            return {"success": False, "error": "Could not extract text from the uploaded PDF"}, 422

        if not os.path.exists(COMPLIANCE_REPORT_PATH):
            return {
                "success": False,
                "error": f"compliance_report.pdf not found at {COMPLIANCE_REPORT_PATH}. "
                         "Please place the file there and restart the server."
            }, 503

        # Extract policy fields for extra LLM context (best-effort)
        policy_dict = extract_insights_with_llm(doc_text)
        if "error" in policy_dict:
            policy_dict = {}

        # Run compliance check against compliance_report.pdf using this document's text
        compliance = check_compliance_with_llm(doc_text, policy_dict=policy_dict)

        return {
            "success":    True,
            "filename":   fname,
            "compliance": compliance,
        }, 200


@ns.route("/renewal-insights")
class RenewalInsights(Resource):
    @ns.doc("renewal_insights")
    @ns.expect(renewal_request_model, validate=True)
    @ns.response(200, "Success", renewal_response_model)
    def post(self):
        """
        Compare a policy against live market quotes.
        Returns actionable insights + side-by-side comparison table.
        Delegates to generate_renewal_insights() from extractor.py.
        """
        data           = request.json
        current_policy = data.get("current_policy", {})
        market_quotes  = data.get("market_quotes", [])
        user_city      = data.get("user_city", "Bengaluru")

        # Calls the tested generate_renewal_insights() from extractor.py
        insights   = generate_renewal_insights(current_policy, market_quotes, user_city)
        best_quote = (
            min(market_quotes, key=lambda q: q.get("annual_premium_inr", 1e9))
            if market_quotes else {}
        )

        def _f(obj, key):
            v = obj.get(key, "-")
            return ", ".join(map(str, v)) if isinstance(v, list) else str(v or "-")

        comparison = [
            ["Field",            "Current Policy",                           "Best Market Quote"],
            ["Plan / Type",      _f(current_policy, "policy_type"),          _f(best_quote, "plan_name")],
            ["Premium",          _f(current_policy, "total_premium"),        str(best_quote.get("annual_premium_inr", "-"))],
            ["IDV",              _f(current_policy, "idv"),                  _f(best_quote, "idv")],
            ["Add-ons",          _f(current_policy, "add_ons"),              _f(best_quote, "key_benefits")],
            ["Coverage",         _f(current_policy, "coverage_details"),     _f(best_quote, "key_benefits")],
            ["Cashless Garages", "-",                                        str(best_quote.get("cashless_garages", "-"))],
        ]

        return {"insights": insights, "comparison": comparison}, 200


# Run
if __name__ == "__main__":
    mongo_uri = os.environ.get("MONGO_URI", "mongodb://localhost:27017")
    mongo_db  = os.environ.get("MONGO_DB",  "insurebot")
    print("\n" + "=" * 60)
    print("  Extractor API")
    print("=" * 60)
    print(f"  Swagger UI : http://localhost:5050/swagger")
    print(f"  MongoDB    : {mongo_uri}")
    print(f"  Database   : {mongo_db}")
    print("=" * 60 + "\n")
    app.run(host="127.0.0.1", port=5050, debug=True, threaded=True)