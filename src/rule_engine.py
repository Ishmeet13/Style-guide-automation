"""
Rule Engine Module
Loads, caches, and manages style guide rules for document validation
"""

import json
import logging
from typing import Dict, List, Optional, Any
from datetime import datetime, timedelta
from pathlib import Path


class RuleEngine:
    """
    Rule Engine for managing document style guide rules
    
    Features:
    - Load rules from JSON files
    - Cache rules in memory
    - Filter rules by category, severity, or enabled status
    - Validate rule structure
    """
    
    def __init__(self, rules_file: str = None, cache_ttl: int = 3600):
        """
        Initialize Rule Engine
        
        Args:
            rules_file: Path to rules JSON file
            cache_ttl: Time to live for cached rules in seconds (default: 1 hour)
        """
        self.logger = logging.getLogger(__name__)
        self.rules_file = rules_file
        self.cache_ttl = cache_ttl
        self.rules: Dict[str, Any] = {}
        self.rules_by_category: Dict[str, List[Dict]] = {}
        self.cache_timestamp: Optional[datetime] = None
        
        if rules_file:
            self.load_rules(rules_file)
    
    def load_rules(self, rules_file: str) -> bool:
        """
        Load rules from JSON file
        
        Args:
            rules_file: Path to rules JSON file
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            rules_path = Path(rules_file)
            
            if not rules_path.exists():
                self.logger.error(f"Rules file not found: {rules_file}")
                return False
            
            with open(rules_path, 'r', encoding='utf-8') as f:
                self.rules = json.load(f)
            
            # Validate rules structure
            if not self._validate_rules_structure():
                self.logger.error("Invalid rules structure")
                return False
            
            # Organize rules by category
            self._organize_rules_by_category()
            
            # Set cache timestamp
            self.cache_timestamp = datetime.now()
            
            self.logger.info(f"Loaded {len(self.rules.get('rules', []))} rules from {rules_file}")
            return True
            
        except json.JSONDecodeError as e:
            self.logger.error(f"Invalid JSON in rules file: {e}")
            return False
        except Exception as e:
            self.logger.error(f"Error loading rules: {e}")
            return False
    
    def _validate_rules_structure(self) -> bool:
        """Validate that rules JSON has required structure"""
        required_fields = ['version', 'rules']
        
        for field in required_fields:
            if field not in self.rules:
                self.logger.error(f"Missing required field: {field}")
                return False
        
        # Validate each rule has required fields
        for rule in self.rules.get('rules', []):
            required_rule_fields = ['rule_id', 'category', 'validation', 'correction_action']
            for field in required_rule_fields:
                if field not in rule:
                    self.logger.error(f"Rule missing required field: {field}")
                    return False
        
        return True
    
    def _organize_rules_by_category(self):
        """Organize rules by category for efficient filtering"""
        self.rules_by_category = {}
        
        for rule in self.rules.get('rules', []):
            category = rule.get('category', 'unknown')
            if category not in self.rules_by_category:
                self.rules_by_category[category] = []
            self.rules_by_category[category].append(rule)
    
    def is_cache_valid(self) -> bool:
        """Check if cached rules are still valid"""
        if not self.cache_timestamp:
            return False
        
        age = datetime.now() - self.cache_timestamp
        return age.total_seconds() < self.cache_ttl
    
    def get_all_rules(self, force_reload: bool = False) -> List[Dict]:
        """
        Get all rules
        
        Args:
            force_reload: Force reload from file even if cache is valid
            
        Returns:
            List of all rules
        """
        if force_reload or not self.is_cache_valid():
            if self.rules_file:
                self.load_rules(self.rules_file)
        
        return self.rules.get('rules', [])
    
    def get_rules_by_category(self, category: str) -> List[Dict]:
        """
        Get all rules for a specific category
        
        Args:
            category: Category name (e.g., 'cover_page', 'tables')
            
        Returns:
            List of rules in that category
        """
        return self.rules_by_category.get(category, [])
    
    def get_rules_by_severity(self, severity: str) -> List[Dict]:
        """
        Get all rules with specific severity
        
        Args:
            severity: Severity level ('high', 'medium', 'low')
            
        Returns:
            List of rules with that severity
        """
        return [
            rule for rule in self.rules.get('rules', [])
            if rule.get('severity') == severity
        ]
    
    def get_enabled_rules(self) -> List[Dict]:
        """Get only enabled rules"""
        return [
            rule for rule in self.rules.get('rules', [])
            if rule.get('enabled', True)
        ]
    
    def get_rule_by_id(self, rule_id: str) -> Optional[Dict]:
        """
        Get a specific rule by ID
        
        Args:
            rule_id: Rule identifier
            
        Returns:
            Rule dictionary or None if not found
        """
        for rule in self.rules.get('rules', []):
            if rule.get('rule_id') == rule_id:
                return rule
        return None
    
    def get_categories(self) -> Dict[str, Any]:
        """Get all category definitions"""
        return self.rules.get('categories', {})
    
    def get_metadata(self) -> Dict[str, Any]:
        """Get rules metadata"""
        return self.rules.get('metadata', {})
    
    def enable_rule(self, rule_id: str) -> bool:
        """Enable a specific rule"""
        rule = self.get_rule_by_id(rule_id)
        if rule:
            rule['enabled'] = True
            return True
        return False
    
    def disable_rule(self, rule_id: str) -> bool:
        """Disable a specific rule"""
        rule = self.get_rule_by_id(rule_id)
        if rule:
            rule['enabled'] = False
            return True
        return False
    
    def get_rules_count(self) -> Dict[str, int]:
        """Get statistics about rules"""
        all_rules = self.rules.get('rules', [])
        return {
            'total': len(all_rules),
            'enabled': len([r for r in all_rules if r.get('enabled', True)]),
            'disabled': len([r for r in all_rules if not r.get('enabled', True)]),
            'high_severity': len([r for r in all_rules if r.get('severity') == 'high']),
            'medium_severity': len([r for r in all_rules if r.get('severity') == 'medium']),
            'low_severity': len([r for r in all_rules if r.get('severity') == 'low']),
            'categories': len(self.rules_by_category)
        }
    
    def __repr__(self) -> str:
        """String representation"""
        stats = self.get_rules_count()
        return (f"RuleEngine(rules={stats['total']}, "
                f"enabled={stats['enabled']}, "
                f"categories={stats['categories']})")


# Example usage
if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    
    # Initialize rule engine
    engine = RuleEngine("bestco-rules.json")
    
    # Get all rules
    print(f"Total rules: {len(engine.get_all_rules())}")
    
    # Get rules by category
    cover_page_rules = engine.get_rules_by_category('cover_page')
    print(f"Cover page rules: {len(cover_page_rules)}")
    
    # Get high severity rules
    high_severity = engine.get_rules_by_severity('high')
    print(f"High severity rules: {len(high_severity)}")
    
    # Print statistics
    print(f"\nRule Engine: {engine}")
    print(f"Statistics: {engine.get_rules_count()}")
