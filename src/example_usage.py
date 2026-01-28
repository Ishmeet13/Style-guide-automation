"""
Complete Usage Example
Demonstrates how to use the Style Guide Automation system
"""

import json
from main import StyleGuideAutomation

def example_1_simple_validation():
    """Example 1: Simple document validation"""
    print("\n" + "="*80)
    print("EXAMPLE 1: Simple Document Validation")
    print("="*80 + "\n")
    
    # Initialize
    automation = StyleGuideAutomation(
        rules_file="bestco-rules.json",
        log_level="INFO"
    )
    
    # Validate document
    report = automation.validate_document("bestco-sample-input.docx")
    
    # Display results
    print(f"Document: {report['document_name']}")
    print(f"Total Violations: {report['summary']['total_violations']}")
    print(f"Processing Time: {report['summary']['processing_duration_seconds']:.2f}s")
    
    print(f"\nViolations by Severity:")
    print(f"  High: {report['summary']['high_severity']}")
    print(f"  Medium: {report['summary']['medium_severity']}")
    print(f"  Low: {report['summary']['low_severity']}")
    
    print(f"\nViolations by Category:")
    for category, count in report['violations_by_category'].items():
        if count > 0:
            print(f"  {category}: {count}")
    
    # Show first 3 violations
    print(f"\nFirst 3 Violations:")
    for i, v in enumerate(report['violations'][:3], 1):
        print(f"\n{i}. [{v['severity'].upper()}] {v['rule_name']}")
        print(f"   Location: Paragraph {v['location'].get('paragraph', 'N/A')}")
        print(f"   Message: {v['message']}")


def example_2_validate_and_correct():
    """Example 2: Validate and correct document"""
    print("\n" + "="*80)
    print("EXAMPLE 2: Validate and Correct Document")
    print("="*80 + "\n")
    
    # Initialize
    automation = StyleGuideAutomation(
        rules_file="bestco-rules.json",
        log_level="INFO"
    )
    
    # Process document
    report = automation.correct_document(
        document_path="bestco-sample-input.docx",
        output_path="bestco-sample-corrected.docx"
    )
    
    # Display results
    summary = report['summary']
    print(f"Document: {report['document_name']}")
    print(f"Status: {report['status']}")
    
    print(f"\nSummary:")
    print(f"  Total Violations: {summary['total_violations']}")
    print(f"  Corrections Applied: {summary['corrections_applied']}")
    print(f"  Corrections Failed: {summary['corrections_failed']}")
    print(f"  Corrections Skipped: {summary['corrections_skipped']}")
    print(f"  Processing Time: {summary['processing_duration_seconds']:.2f}s")
    
    # Show corrections by category
    print(f"\nCorrections by Category:")
    for category, stats in report['corrections_summary']['by_category'].items():
        if stats['violations'] > 0:
            success_rate = (stats['corrections_applied'] / stats['violations'] * 100)
            print(f"  {category:25}: {stats['violations']:2} violations, "
                  f"{stats['corrections_applied']:2} fixed ({success_rate:.0f}%)")
    
    # Show failed corrections
    if report['failed_corrections']:
        print(f"\nFailed Corrections (Require Manual Review):")
        for failed in report['failed_corrections']:
            print(f"  - {failed['rule_id']}: {failed['reason']}")
    
    # Save report
    automation.save_report(report, "bestco-validation-report.json")
    print(f"\nFiles Generated:")
    print(f"  - Corrected Document: bestco-sample-corrected.docx")
    print(f"  - Validation Report: bestco-validation-report.json")


def example_3_advanced_usage():
    """Example 3: Advanced usage with individual components"""
    print("\n" + "="*80)
    print("EXAMPLE 3: Advanced Usage with Individual Components")
    print("="*80 + "\n")
    
    from rule_engine import RuleEngine
    from validator import DocumentValidator
    from corrector import DocumentCorrector
    
    # Step 1: Load rules
    print("Step 1: Loading rules...")
    engine = RuleEngine("bestco-rules.json")
    
    rule_stats = engine.get_rules_count()
    print(f"  Loaded {rule_stats['total']} rules")
    print(f"  Enabled: {rule_stats['enabled']}")
    print(f"  Categories: {rule_stats['categories']}")
    
    # Step 2: Get specific rule categories
    print("\nStep 2: Filtering rules...")
    cover_page_rules = engine.get_rules_by_category('cover_page')
    print(f"  Cover page rules: {len(cover_page_rules)}")
    
    high_severity = engine.get_rules_by_severity('high')
    print(f"  High severity rules: {len(high_severity)}")
    
    # Step 3: Validate document
    print("\nStep 3: Validating document...")
    validator = DocumentValidator(engine)
    violations = validator.validate_document("bestco-sample-input.docx")
    print(f"  Found {len(violations)} violations")
    
    # Step 4: Filter violations
    print("\nStep 4: Analyzing violations...")
    high_sev_violations = validator.get_violations_by_severity('high')
    print(f"  High severity violations: {len(high_sev_violations)}")
    
    cover_violations = validator.get_violations_by_category('cover_page')
    print(f"  Cover page violations: {len(cover_violations)}")
    
    # Step 5: Apply corrections
    print("\nStep 5: Applying corrections...")
    corrector = DocumentCorrector(engine)
    results = corrector.apply_corrections(
        document_path="bestco-sample-input.docx",
        violations=violations,
        output_path="bestco-sample-corrected-advanced.docx"
    )
    
    stats = corrector.get_correction_stats()
    print(f"  Corrections applied: {stats['applied']}")
    print(f"  Corrections failed: {stats['failed']}")


def example_4_rule_management():
    """Example 4: Rule management and filtering"""
    print("\n" + "="*80)
    print("EXAMPLE 4: Rule Management and Filtering")
    print("="*80 + "\n")
    
    from rule_engine import RuleEngine
    
    engine = RuleEngine("bestco-rules.json")
    
    # List all categories
    categories = engine.get_categories()
    print("Available Categories:")
    for cat_name, cat_info in categories.items():
        print(f"  - {cat_name}: {cat_info.get('name')}")
    
    # Get specific rule
    print(f"\nSpecific Rule Details:")
    rule = engine.get_rule_by_id("COVER_PAGE_CENTER_ALIGNMENT")
    if rule:
        print(f"  Rule ID: {rule['rule_id']}")
        print(f"  Category: {rule['category']}")
        print(f"  Priority: {rule['priority']}")
        print(f"  Severity: {rule['severity']}")
        print(f"  Enabled: {rule['enabled']}")
        print(f"  Description: {rule['description']}")
    
    # Get metadata
    print(f"\nRules Metadata:")
    metadata = engine.get_metadata()
    print(f"  Version: {metadata.get('version')}")
    print(f"  Organization: {metadata.get('organization')}")
    print(f"  Total Rules: {metadata.get('total_rules')}")
    print(f"  Active Rules: {metadata.get('active_rules')}")


def example_5_error_handling():
    """Example 5: Error handling and robustness"""
    print("\n" + "="*80)
    print("EXAMPLE 5: Error Handling")
    print("="*80 + "\n")
    
    from main import StyleGuideAutomation
    
    try:
        # Try with non-existent file
        automation = StyleGuideAutomation(
            rules_file="bestco-rules.json",
            log_level="WARNING"
        )
        
        report = automation.validate_document("nonexistent.docx")
        
    except FileNotFoundError as e:
        print(f"✓ Caught FileNotFoundError: {e}")
    
    try:
        # Try with invalid rules file
        automation = StyleGuideAutomation(
            rules_file="invalid-rules.json",
            log_level="WARNING"
        )
        
    except Exception as e:
        print(f"✓ Caught Exception for invalid rules: {type(e).__name__}")
    
    print("\n✓ Error handling working correctly")


def example_6_performance_test():
    """Example 6: Performance testing"""
    print("\n" + "="*80)
    print("EXAMPLE 6: Performance Testing")
    print("="*80 + "\n")
    
    import time
    from main import StyleGuideAutomation
    
    automation = StyleGuideAutomation(
        rules_file="bestco-rules.json",
        log_level="WARNING"
    )
    
    # Test 1: Validation performance
    print("Test 1: Validation Performance")
    start = time.time()
    report = automation.validate_document("bestco-sample-input.docx")
    elapsed = time.time() - start
    print(f"  Time: {elapsed:.2f}s")
    print(f"  Violations: {report['summary']['total_violations']}")
    print(f"  Violations/second: {report['summary']['total_violations']/elapsed:.1f}")
    
    # Test 2: Correction performance
    print(f"\nTest 2: Correction Performance")
    start = time.time()
    report = automation.correct_document(
        "bestco-sample-input.docx",
        "bestco-sample-corrected-perf.docx"
    )
    elapsed = time.time() - start
    print(f"  Time: {elapsed:.2f}s")
    print(f"  Corrections: {report['summary']['corrections_applied']}")
    print(f"  Corrections/second: {report['summary']['corrections_applied']/elapsed:.1f}")


def main():
    """Run all examples"""
    print("\n")
    print("="*80)
    print(" "*20 + "STYLE GUIDE AUTOMATION")
    print(" "*25 + "Usage Examples")
    print("="*80)
    
    # Run examples
    example_1_simple_validation()
    example_2_validate_and_correct()
    example_3_advanced_usage()
    example_4_rule_management()
    example_5_error_handling()
    example_6_performance_test()
    
    print("\n" + "="*80)
    print("All examples completed successfully!")
    print("="*80 + "\n")


if __name__ == "__main__":
    main()
