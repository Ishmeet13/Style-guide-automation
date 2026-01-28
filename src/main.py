"""
Main Orchestrator
Coordinates Rule Engine, Validator, and Corrector for complete document processing
"""

import json
import logging
from typing import Dict, Any, Optional
from datetime import datetime
from pathlib import Path

from rule_engine import RuleEngine
from validator import DocumentValidator
from corrector import DocumentCorrector


class StyleGuideAutomation:
    """
    Main orchestrator for style guide automation
    
    Provides high-level API for validating and correcting documents
    """
    
    def __init__(self, rules_file: str, cache_ttl: int = 3600, log_level: str = 'INFO'):
        """
        Initialize Style Guide Automation
        
        Args:
            rules_file: Path to rules JSON file
            cache_ttl: Cache TTL in seconds (default: 1 hour)
            log_level: Logging level (default: INFO)
        """
        # Setup logging
        logging.basicConfig(
            level=getattr(logging, log_level.upper()),
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger(__name__)
        
        # Initialize components
        self.rule_engine = RuleEngine(rules_file, cache_ttl)
        self.validator = DocumentValidator(self.rule_engine)
        self.corrector = DocumentCorrector(self.rule_engine)
        
        self.logger.info("StyleGuideAutomation initialized")
    
    def validate_document(self, document_path: str) -> Dict[str, Any]:
        """
        Validate a document and return report
        
        Args:
            document_path: Path to .docx document
            
        Returns:
            Validation report dictionary
        """
        self.logger.info(f"Validating document: {document_path}")
        start_time = datetime.now()
        
        try:
            # Validate
            violations = self.validator.validate_document(document_path)
            
            # Generate report
            end_time = datetime.now()
            processing_time = (end_time - start_time).total_seconds()
            
            report = self._generate_validation_report(
                document_path=document_path,
                violations=violations,
                processing_time=processing_time
            )
            
            return report
            
        except Exception as e:
            self.logger.error(f"Error validating document: {e}")
            raise
    
    def correct_document(self, document_path: str, output_path: str) -> Dict[str, Any]:
        """
        Validate and correct a document
        
        Args:
            document_path: Path to input document
            output_path: Path for corrected document
            
        Returns:
            Complete report with corrections
        """
        self.logger.info(f"Processing document: {document_path}")
        start_time = datetime.now()
        
        try:
            # Step 1: Validate
            violations = self.validator.validate_document(document_path)
            
            # Step 2: Apply corrections
            correction_results = []
            if violations:
                correction_results = self.corrector.apply_corrections(
                    document_path=document_path,
                    violations=violations,
                    output_path=output_path
                )
            
            # Generate complete report
            end_time = datetime.now()
            processing_time = (end_time - start_time).total_seconds()
            
            report = self._generate_complete_report(
                document_path=document_path,
                output_path=output_path,
                violations=violations,
                correction_results=correction_results,
                processing_time=processing_time
            )
            
            return report
            
        except Exception as e:
            self.logger.error(f"Error processing document: {e}")
            raise
    
    def _generate_validation_report(self, document_path: str, violations: list,
                                   processing_time: float) -> Dict[str, Any]:
        """Generate validation-only report"""
        summary = self.validator.get_summary()
        
        return {
            'job_id': self._generate_job_id(),
            'document_name': Path(document_path).name,
            'processing_timestamp': datetime.now().isoformat(),
            'status': 'validation_complete',
            'summary': {
                'total_violations': summary['total_violations'],
                'high_severity': summary['high_severity'],
                'medium_severity': summary['medium_severity'],
                'low_severity': summary['low_severity'],
                'processing_duration_seconds': round(processing_time, 2),
                'rules_checked': len(self.rule_engine.get_enabled_rules())
            },
            'violations': [v.to_dict() for v in violations],
            'violations_by_category': summary['by_category'],
            'document_info': {
                'original_file': Path(document_path).name,
                'original_file_path': document_path
            }
        }
    
    def _generate_complete_report(self, document_path: str, output_path: str,
                                 violations: list, correction_results: list,
                                 processing_time: float) -> Dict[str, Any]:
        """Generate complete report with corrections"""
        summary = self.validator.get_summary()
        correction_stats = self.corrector.get_correction_stats()
        
        # Calculate corrections by category
        corrections_by_category = {}
        for category in self.rule_engine.rules_by_category.keys():
            category_violations = [
                v for v in violations
                if self.rule_engine.get_rule_by_id(v.rule_id).get('category') == category
            ]
            
            corrections_by_category[category] = {
                'violations': len(category_violations),
                'corrections_applied': len([
                    v for v in category_violations
                    if v.correction_status == 'applied'
                ]),
                'corrections_failed': len([
                    v for v in category_violations
                    if v.correction_status == 'failed'
                ])
            }
        
        # Calculate corrections by severity
        corrections_by_severity = {}
        for severity in ['high', 'medium', 'low']:
            sev_violations = [v for v in violations if v.severity == severity]
            
            corrections_by_severity[severity] = {
                'violations': len(sev_violations),
                'corrections_applied': len([
                    v for v in sev_violations
                    if v.correction_status == 'applied'
                ]),
                'corrections_failed': len([
                    v for v in sev_violations
                    if v.correction_status == 'failed'
                ])
            }
        
        # Failed corrections
        failed_corrections = [
            {
                'violation_id': r.violation_id,
                'rule_id': r.rule_id,
                'reason': r.error_details or 'Unknown error'
            }
            for r in correction_results if r.status == 'failed'
        ]
        
        return {
            'job_id': self._generate_job_id(),
            'document_name': Path(document_path).name,
            'processing_timestamp': datetime.now().isoformat(),
            'status': 'completed' if not failed_corrections else 'partial',
            'summary': {
                'total_violations': summary['total_violations'],
                'corrections_applied': correction_stats['applied'],
                'corrections_failed': correction_stats['failed'],
                'corrections_skipped': correction_stats['skipped'],
                'rules_checked': len(self.rule_engine.get_enabled_rules()),
                'processing_duration_seconds': round(processing_time, 2)
            },
            'violations': [v.to_dict() for v in violations],
            'corrections_summary': {
                'by_category': corrections_by_category,
                'by_severity': corrections_by_severity
            },
            'failed_corrections': failed_corrections,
            'document_info': {
                'original_file': Path(document_path).name,
                'corrected_file': Path(output_path).name,
                'original_file_path': document_path,
                'corrected_file_path': output_path
            },
            'execution_metadata': {
                'rules_version': self.rule_engine.rules.get('version'),
                'total_rules': len(self.rule_engine.rules.get('rules', [])),
                'enabled_rules': len(self.rule_engine.get_enabled_rules())
            },
            'recommendations': self._generate_recommendations(failed_corrections),
            'next_steps': self._generate_next_steps(output_path, failed_corrections)
        }
    
    def _generate_job_id(self) -> str:
        """Generate unique job ID"""
        import uuid
        return str(uuid.uuid4())[:8]
    
    def _generate_recommendations(self, failed_corrections: list) -> list:
        """Generate recommendations based on results"""
        recommendations = []
        
        if failed_corrections:
            recommendations.append(
                f"Review {len(failed_corrections)} failed corrections manually"
            )
            recommendations.append(
                "Check if failed corrections require manual intervention"
            )
        
        return recommendations
    
    def _generate_next_steps(self, output_path: str, failed_corrections: list) -> list:
        """Generate next steps"""
        steps = [
            f"Review corrected document: {output_path}"
        ]
        
        if failed_corrections:
            steps.append(f"Manually fix {len(failed_corrections)} failed corrections")
        
        return steps
    
    def save_report(self, report: Dict[str, Any], output_path: str):
        """Save report to JSON file"""
        try:
            with open(output_path, 'w') as f:
                json.dump(report, f, indent=2, default=str)
            self.logger.info(f"Report saved to: {output_path}")
        except Exception as e:
            self.logger.error(f"Error saving report: {e}")
            raise


def main():
    """Main entry point for CLI usage"""
    import argparse
    
    parser = argparse.ArgumentParser(
        description='Style Guide Automation - Validate and correct document formatting'
    )
    parser.add_argument('document', help='Path to document to process')
    parser.add_argument('--rules', required=True, help='Path to rules JSON file')
    parser.add_argument('--output', help='Path for corrected document')
    parser.add_argument('--report', help='Path to save JSON report')
    parser.add_argument('--validate-only', action='store_true',
                       help='Only validate, do not apply corrections')
    parser.add_argument('--log-level', default='INFO',
                       choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
                       help='Logging level')
    
    args = parser.parse_args()
    
    # Initialize automation
    automation = StyleGuideAutomation(
        rules_file=args.rules,
        log_level=args.log_level
    )
    
    # Process document
    if args.validate_only:
        # Validate only
        report = automation.validate_document(args.document)
        print(f"\n{'='*80}")
        print("VALIDATION REPORT")
        print(f"{'='*80}")
    else:
        # Validate and correct
        if not args.output:
            # Generate output filename
            doc_path = Path(args.document)
            args.output = str(doc_path.parent / f"{doc_path.stem}_corrected{doc_path.suffix}")
        
        report = automation.correct_document(args.document, args.output)
        print(f"\n{'='*80}")
        print("CORRECTION REPORT")
        print(f"{'='*80}")
    
    # Print summary
    summary = report['summary']
    print(f"\nDocument: {report['document_name']}")
    print(f"Status: {report['status']}")
    print(f"\nSummary:")
    print(f"  Total Violations: {summary['total_violations']}")
    
    if 'corrections_applied' in summary:
        print(f"  Corrections Applied: {summary['corrections_applied']}")
        print(f"  Corrections Failed: {summary['corrections_failed']}")
    
    print(f"  Processing Time: {summary['processing_duration_seconds']:.2f}s")
    
    # Print violations by category
    if 'violations_by_category' in report:
        print(f"\nViolations by Category:")
        for category, count in report['violations_by_category'].items():
            if count > 0:
                print(f"  {category}: {count}")
    
    # Print failed corrections
    if report.get('failed_corrections'):
        print(f"\nFailed Corrections (require manual review):")
        for failed in report['failed_corrections']:
            print(f"  - {failed['rule_id']}: {failed['reason']}")
    
    # Print recommendations
    if report.get('recommendations'):
        print(f"\nRecommendations:")
        for i, rec in enumerate(report['recommendations'], 1):
            print(f"  {i}. {rec}")
    
    # Print next steps
    if report.get('next_steps'):
        print(f"\nNext Steps:")
        for i, step in enumerate(report['next_steps'], 1):
            print(f"  {i}. {step}")
    
    print(f"\n{'='*80}\n")
    
    # Save report if requested
    if args.report:
        automation.save_report(report, args.report)
        print(f"Report saved to: {args.report}")


if __name__ == "__main__":
    main()
