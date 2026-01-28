"""
Corrector Module
Applies corrections to documents based on detected violations
"""

import logging
from typing import List, Dict, Any
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH


class CorrectionResult:
    """Represents the result of applying a correction"""
    
    def __init__(self, violation_id: int, rule_id: str, status: str, message: str):
        self.violation_id = violation_id
        self.rule_id = rule_id
        self.status = status  # 'applied', 'failed', 'skipped'
        self.message = message
        self.timestamp = datetime.now()
        self.error_details = None
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary"""
        return {
            'violation_id': self.violation_id,
            'rule_id': self.rule_id,
            'status': self.status,
            'message': self.message,
            'timestamp': self.timestamp.isoformat(),
            'error_details': self.error_details
        }


class DocumentCorrector:
    """
    Document Corrector
    
    Applies corrections to .docx documents based on violations
    Preserves document content while fixing formatting
    """
    
    def __init__(self, rule_engine):
        """
        Initialize Corrector
        
        Args:
            rule_engine: RuleEngine instance with loaded rules
        """
        self.logger = logging.getLogger(__name__)
        self.rule_engine = rule_engine
        self.correction_results: List[CorrectionResult] = []
    
    def apply_corrections(self, document_path: str, violations: List, output_path: str) -> List[CorrectionResult]:
        """
        Apply corrections to a document
        
        Args:
            document_path: Path to input document
            violations: List of Violation objects
            output_path: Path for corrected document
            
        Returns:
            List of correction results
        """
        self.logger.info(f"Applying corrections to: {document_path}")
        self.logger.info(f"Processing {len(violations)} violations")
        
        self.correction_results = []
        
        try:
            # Load document
            doc = Document(document_path)
            
            # Sort violations by priority (via rule priority)
            sorted_violations = sorted(
                violations,
                key=lambda v: self._get_rule_priority(v.rule_id)
            )
            
            # Apply each correction
            for violation in sorted_violations:
                try:
                    result = self._apply_correction(doc, violation)
                    self.correction_results.append(result)
                    
                    # Update violation status
                    violation.correction_status = result.status
                    violation.correction_timestamp = result.timestamp
                    
                except Exception as e:
                    self.logger.error(f"Error correcting {violation.rule_id}: {e}")
                    result = CorrectionResult(
                        violation_id=violation.violation_id,
                        rule_id=violation.rule_id,
                        status='failed',
                        message=f"Correction failed: {str(e)}"
                    )
                    result.error_details = str(e)
                    self.correction_results.append(result)
            
            # Save corrected document
            doc.save(output_path)
            self.logger.info(f"Corrected document saved to: {output_path}")
            
            # Log statistics
            stats = self.get_correction_stats()
            self.logger.info(f"Corrections applied: {stats['applied']}, failed: {stats['failed']}")
            
            return self.correction_results
            
        except Exception as e:
            self.logger.error(f"Error in correction process: {e}")
            raise
    
    def _apply_correction(self, doc: Document, violation) -> CorrectionResult:
        """Apply a single correction"""
        rule = self.rule_engine.get_rule_by_id(violation.rule_id)
        
        if not rule:
            return CorrectionResult(
                violation_id=violation.violation_id,
                rule_id=violation.rule_id,
                status='failed',
                message='Rule not found'
            )
        
        action = rule.get('correction_action', {})
        action_type = action.get('type')
        
        try:
            # Route to appropriate correction method
            if action_type == 'apply_formatting':
                self._apply_formatting(doc, violation, action)
            elif action_type == 'apply_alignment':
                self._apply_alignment(doc, violation, action)
            elif action_type == 'apply_cover_page_company_formatting':
                self._apply_formatting(doc, violation, action)  # Same as apply_formatting
            elif action_type == 'ensure_blank_row':
                self._ensure_blank_row(doc, violation, action)
            elif action_type == 'apply_table_row_height':
                self._apply_table_row_height(doc, violation, action)
            elif action_type == 'apply_column_width':
                self._apply_column_width(doc, violation, action)
            elif action_type == 'apply_bold_to_current_period':
                self._apply_bold_to_current_period(doc, violation, action)
            else:
                return CorrectionResult(
                    violation_id=violation.violation_id,
                    rule_id=violation.rule_id,
                    status='skipped',
                    message=f'Unknown action type: {action_type}'
                )
            
            return CorrectionResult(
                violation_id=violation.violation_id,
                rule_id=violation.rule_id,
                status='applied',
                message=f'Successfully applied {action_type}'
            )
            
        except Exception as e:
            self.logger.error(f"Error applying {action_type}: {e}")
            result = CorrectionResult(
                violation_id=violation.violation_id,
                rule_id=violation.rule_id,
                status='failed',
                message=f'Failed to apply {action_type}'
            )
            result.error_details = str(e)
            return result
    
    def _apply_formatting(self, doc: Document, violation, action: Dict[str, Any]):
        """Apply general formatting correction"""
        para_index = violation.location.get('paragraph')
        
        if para_index is None or para_index >= len(doc.paragraphs):
            raise ValueError(f"Invalid paragraph index: {para_index}")
        
        para = doc.paragraphs[para_index]
        props = action.get('properties', {})
        
        # Apply alignment
        if 'alignment' in props:
            para.alignment = self._parse_alignment(props['alignment'])
        
        # Apply font properties to all runs
        for run in para.runs:
            if 'font_name' in props:
                run.font.name = props['font_name']
            
            if 'font_size' in props:
                run.font.size = Pt(props['font_size'])
            
            if 'bold' in props:
                run.font.bold = props['bold']
            
            if 'italic' in props:
                run.font.italic = props.get('italic')
        
        self.logger.debug(f"Applied formatting to paragraph {para_index}")
    
    def _apply_alignment(self, doc: Document, violation, action: Dict[str, Any]):
        """Apply alignment correction"""
        para_index = violation.location.get('paragraph')
        
        if para_index is None or para_index >= len(doc.paragraphs):
            raise ValueError(f"Invalid paragraph index: {para_index}")
        
        para = doc.paragraphs[para_index]
        props = action.get('properties', {})
        
        if 'alignment' in props:
            para.alignment = self._parse_alignment(props['alignment'])
            self.logger.debug(f"Applied alignment {props['alignment']} to paragraph {para_index}")
    
    def _ensure_blank_row(self, doc: Document, violation, action: Dict[str, Any]):
        """Ensure a row is blank"""
        para_index = violation.location.get('paragraph')
        
        if para_index is None or para_index >= len(doc.paragraphs):
            raise ValueError(f"Invalid paragraph index: {para_index}")
        
        para = doc.paragraphs[para_index]
        
        # Clear text
        para.clear()
        
        # Set font properties
        props = action.get('properties', {})
        run = para.add_run()
        
        if 'font_name' in props:
            run.font.name = props['font_name']
        if 'font_size' in props:
            run.font.size = Pt(props['font_size'])
        
        self.logger.debug(f"Ensured blank row at paragraph {para_index}")
    
    def _apply_table_row_height(self, doc: Document, violation, action: Dict[str, Any]):
        """Apply table row height correction"""
        table_idx = violation.location.get('table')
        row_idx = violation.location.get('row')
        
        if table_idx is None or row_idx is None:
            raise ValueError("Missing table or row index")
        
        if table_idx >= len(doc.tables):
            raise ValueError(f"Invalid table index: {table_idx}")
        
        table = doc.tables[table_idx]
        
        if row_idx >= len(table.rows):
            raise ValueError(f"Invalid row index: {row_idx}")
        
        row = table.rows[row_idx]
        props = action.get('properties', {})
        
        if 'row_height' in props:
            height_cm = props['row_height']
            # Convert cm to inches (1 cm = 0.393701 inches)
            row.height = Inches(height_cm / 2.54)
        
        self.logger.debug(f"Applied row height to table {table_idx}, row {row_idx}")
    
    def _apply_column_width(self, doc: Document, violation, action: Dict[str, Any]):
        """Apply column width correction"""
        table_idx = violation.location.get('table')
        col_idx = violation.location.get('column')
        
        if table_idx is None or col_idx is None:
            raise ValueError("Missing table or column index")
        
        if table_idx >= len(doc.tables):
            raise ValueError(f"Invalid table index: {table_idx}")
        
        table = doc.tables[table_idx]
        
        if col_idx >= len(table.columns):
            raise ValueError(f"Invalid column index: {col_idx}")
        
        col = table.columns[col_idx]
        props = action.get('properties', {})
        
        if 'column_width' in props:
            width_cm = props['column_width']
            col.width = Inches(width_cm / 2.54)
        
        self.logger.debug(f"Applied column width to table {table_idx}, column {col_idx}")
    
    def _apply_bold_to_current_period(self, doc: Document, violation, action: Dict[str, Any]):
        """Apply bold to current period column"""
        table_idx = violation.location.get('table')
        row_idx = violation.location.get('row')
        col_idx = violation.location.get('column')
        
        if None in [table_idx, row_idx, col_idx]:
            raise ValueError("Missing location information")
        
        if table_idx >= len(doc.tables):
            raise ValueError(f"Invalid table index: {table_idx}")
        
        table = doc.tables[table_idx]
        
        if row_idx >= len(table.rows):
            raise ValueError(f"Invalid row index: {row_idx}")
        
        row = table.rows[row_idx]
        
        if col_idx >= len(row.cells):
            raise ValueError(f"Invalid column index: {col_idx}")
        
        cell = row.cells[col_idx]
        
        # Make all runs bold
        for para in cell.paragraphs:
            for run in para.runs:
                run.font.bold = True
        
        self.logger.debug(f"Applied bold to table {table_idx}, row {row_idx}, col {col_idx}")
    
    def _parse_alignment(self, alignment_str: str) -> WD_ALIGN_PARAGRAPH:
        """Convert alignment string to enum"""
        alignment_map = {
            'left': WD_ALIGN_PARAGRAPH.LEFT,
            'center': WD_ALIGN_PARAGRAPH.CENTER,
            'right': WD_ALIGN_PARAGRAPH.RIGHT,
            'justify': WD_ALIGN_PARAGRAPH.JUSTIFY
        }
        return alignment_map.get(alignment_str.lower(), WD_ALIGN_PARAGRAPH.LEFT)
    
    def _get_rule_priority(self, rule_id: str) -> int:
        """Get priority of a rule"""
        rule = self.rule_engine.get_rule_by_id(rule_id)
        return rule.get('priority', 999) if rule else 999
    
    def get_correction_stats(self) -> Dict[str, int]:
        """Get statistics on corrections"""
        return {
            'total': len(self.correction_results),
            'applied': len([r for r in self.correction_results if r.status == 'applied']),
            'failed': len([r for r in self.correction_results if r.status == 'failed']),
            'skipped': len([r for r in self.correction_results if r.status == 'skipped'])
        }


# Example usage
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    
    from rule_engine import RuleEngine
    from validator import DocumentValidator
    
    # Initialize components
    engine = RuleEngine("bestco-rules.json")
    validator = DocumentValidator(engine)
    corrector = DocumentCorrector(engine)
    
    # Validate document
    violations = validator.validate_document("bestco-sample-input.docx")
    print(f"Found {len(violations)} violations")
    
    # Apply corrections
    if violations:
        results = corrector.apply_corrections(
            document_path="bestco-sample-input.docx",
            violations=violations,
            output_path="bestco-sample-corrected.docx"
        )
        
        # Print results
        stats = corrector.get_correction_stats()
        print(f"\nCorrection Results:")
        print(f"  Applied: {stats['applied']}")
        print(f"  Failed: {stats['failed']}")
        print(f"  Skipped: {stats['skipped']}")
        
        # Show failed corrections
        failed = [r for r in results if r.status == 'failed']
        if failed:
            print(f"\nFailed Corrections:")
            for r in failed:
                print(f"  - {r.rule_id}: {r.message}")
