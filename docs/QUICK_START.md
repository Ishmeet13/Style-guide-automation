# Quick Implementation Guide

**Style Guide Automation - Getting Started**

---

## Quick Start

### Prerequisites

```bash
# Install Python 3.10+
python --version  # Should be 3.10+

# Install dependencies
pip install python-docx fastapi uvicorn celery redis
```

### Create Basic Validator

```python
# validator.py

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

def validate_cover_page(doc_path):
    """Validate and correct cover page formatting"""
    
    # Load document
    doc = Document(doc_path)
    
    violations = []
    
    # Check first 20 paragraphs for company name
    for i, para in enumerate(doc.paragraphs[:20]):
        if 'formerly' in para.text.lower():
            # Found title paragraph
            
            # Check alignment
            if para.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                violations.append({
                    'rule': 'title_alignment',
                    'expected': 'center',
                    'actual': para.alignment,
                    'paragraph_index': i
                })
                
                # Auto-correct
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Check font
            for run in para.runs:
                if run.font.name != 'Arial':
                    violations.append({
                        'rule': 'title_font',
                        'expected': 'Arial',
                        'actual': run.font.name
                    })
                    
                    # Auto-correct
                    run.font.name = 'Arial'
                
                if run.font.size != Pt(14):
                    violations.append({
                        'rule': 'title_size',
                        'expected': '14pt',
                        'actual': str(run.font.size.pt) + 'pt' if run.font.size else 'None'
                    })
                    
                    # Auto-correct
                    run.font.size = Pt(14)
                
                if not run.font.bold:
                    violations.append({
                        'rule': 'title_bold',
                        'expected': True,
                        'actual': False
                    })
                    
                    # Auto-correct
                    run.font.bold = True
    
    # Save corrected document
    output_path = doc_path.replace('.docx', '_corrected.docx')
    doc.save(output_path)
    
    return violations, output_path

# Test it
if __name__ == '__main__':
    violations, output = validate_cover_page('input.docx')
    
    print(f"Found {len(violations)} violations")
    print(f"Corrected document saved to: {output}")
    
    for v in violations:
        print(f"  - {v['rule']}: expected={v['expected']}, actual={v['actual']}")
```

### Run It

```bash
python validator.py

# Output:
# Found 4 violations
# Corrected document saved to: input_corrected.docx
#   - title_alignment: expected=center, actual=0
#   - title_font: expected=Arial, actual=Calibri
#   - title_size: expected=14pt, actual=11.0pt
#   - title_bold: expected=True, actual=False
```

---

## Full System Setup

### Step 1: Clone Repository

```bash
git clone https://github.com/your-org/style-guide-automation.git
cd style-guide-automation
```

### Step 2: Install Dependencies

```bash
# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install requirements
pip install -r requirements.txt
```

### Step 3: Start Infrastructure

```bash
# Start PostgreSQL, Redis, MinIO using Docker Compose
docker-compose up -d
```

### Step 4: Run Database Migrations

```bash
# Initialize database schema
alembic upgrade head
```

### Step 5: Load Rules

```bash
# Load default rule set
python scripts/load_rules.py rules/financial_reports_v1.json
```

### Step 6: Start API Server

```bash
# Start FastAPI server
uvicorn api.app:app --reload --host 0.0.0.0 --port 8000
```

### Step 7: Start Workers

```bash
# In a new terminal
celery -A services.queue worker --loglevel=info
```

### Step 8: Test It

```bash
# Submit a test document
curl -X POST http://localhost:8000/api/v1/documents \
  -H "Authorization: Bearer YOUR_TOKEN" \
  -F "file=@test_input.docx" \
  -F "rule_set=financial_reports_v1"

# Response:
# {
#   "job_id": "job_abc123",
#   "status": "queued"
# }

# Check status
curl http://localhost:8000/api/v1/documents/job_abc123 \
  -H "Authorization: Bearer YOUR_TOKEN"

# Download result
curl http://localhost:8000/api/v1/documents/job_abc123/result \
  -H "Authorization: Bearer YOUR_TOKEN" \
  -o corrected.docx
```

---

## Key Files to Understand

### 1. Rule Definition (`rules/financial_reports_v1.json`)

```json
{
  "rule_set_id": "financial_reports_v1",
  "version": "1.0.0",
  "rules": [
    {
      "rule_id": "cover_page_title_alignment",
      "category": "cover_page",
      "element": "paragraph",
      "selector": {
        "contains_text": "formerly"
      },
      "checks": [
        {
          "property": "alignment",
          "expected": "center",
          "severity": "error"
        }
      ],
      "corrections": [
        {
          "action": "set_alignment",
          "value": "center"
        }
      ]
    }
  ]
}
```

### 2. Validator (`core/validator.py`)

```python
class ValidationEngine:
    def validate_document(self, doc):
        """Main validation entry point"""
        violations = []
        
        for rule in self.rules:
            # Find elements matching rule selector
            elements = self.select_elements(doc, rule.selector)
            
            # Check each element against rule checks
            for element in elements:
                for check in rule.checks:
                    if not self.passes_check(element, check):
                        violations.append(Violation(
                            rule_id=rule.rule_id,
                            element=element,
                            check=check
                        ))
        
        return violations
```

### 3. Corrector (`core/corrector.py`)

```python
class CorrectionEngine:
    def correct_violations(self, doc, violations):
        """Apply corrections for violations"""
        
        for violation in violations:
            rule = self.find_rule(violation.rule_id)
            
            for correction in rule.corrections:
                self.apply_correction(doc, violation, correction)
        
        return doc
```

### 4. API Route (`api/routes/documents.py`)

```python
@app.post("/api/v1/documents")
async def submit_document(
    file: UploadFile,
    rule_set: str = "financial_reports_v1"
):
    # Save file
    job_id = generate_job_id()
    file_path = await storage.save_file(job_id, file)
    
    # Queue processing
    task = process_document.delay(job_id, file_path, rule_set)
    
    return {
        "job_id": job_id,
        "status": "queued"
    }
```

---

## Testing

### Run Unit Tests

```bash
pytest tests/unit/ -v
```

### Run Integration Tests

```bash
pytest tests/integration/ -v
```

### Run All Tests with Coverage

```bash
pytest --cov=core --cov=api --cov-report=html
```

---

## Monitoring

### View Metrics

```bash
# Prometheus metrics endpoint
curl http://localhost:8000/metrics

# Grafana dashboard
open http://localhost:3000
```

### View Logs

```bash
# API logs
docker logs styleguide-api -f

# Worker logs
docker logs styleguide-worker -f
```

---

## Troubleshooting

### Issue: Document won't parse

```python
# Check file validity
from docx import Document

try:
    doc = Document('problem.docx')
    print("Document is valid")
except Exception as e:
    print(f"Document error: {e}")
```

### Issue: Rules not loading

```bash
# Validate rule JSON
python -m json.tool rules/your_rules.json

# Check rule schema
python scripts/validate_rules.py rules/your_rules.json
```

### Issue: Worker not processing

```bash
# Check Redis connection
redis-cli ping

# Check Celery queue
celery -A services.queue inspect active

# Restart worker
docker-compose restart worker
```

---

## Additional Resources

- **Research Document:** See `RESEARCH_DOCUMENT.md` for technology comparison
- **System Design:** See `SYSTEM_DESIGN.md` for complete architecture
- **API Docs:** http://localhost:8000/docs (Swagger UI)
- **python-docx Docs:** https://python-docx.readthedocs.io/

---

## Learning Path

### 1. Start Here (Day 1)
- Read EXECUTIVE_SUMMARY.md
- Run Quick Start example
- Understand basic validation/correction

### 2. Deep Dive (Day 2-3)
- Read SYSTEM_DESIGN.md
- Study rule engine design
- Explore API specification

### 3. Hands-On (Day 4-5)
- Implement custom rules
- Test with real documents
- Deploy to dev environment

### 4. Production Ready (Week 2)
- Set up monitoring
- Configure backups
- Security hardening
- Performance tuning

---

## 💡 Pro Tips

**Tip 1: Cache Rules**
```python
@lru_cache(maxsize=10)
def load_rules(rule_set_id):
    # Rules don't change often - cache them!
    return RuleEngine.load(rule_set_id)
```

**Tip 2: Batch Processing**
```python
# Process multiple documents in parallel
with ThreadPoolExecutor(max_workers=5) as executor:
    futures = [executor.submit(process_doc, doc) for doc in docs]
    results = [f.result() for f in futures]
```

**Tip 3: Progress Tracking**
```python
# Use Celery for progress updates
@celery_app.task(bind=True)
def process_document(self, job_id):
    self.update_state(state='PROGRESS', meta={'current': 25, 'total': 100})
    # ... validation (25%)
    
    self.update_state(state='PROGRESS', meta={'current': 75, 'total': 100})
    # ... correction (75%)
    
    return {'status': 'completed'}
```

---
