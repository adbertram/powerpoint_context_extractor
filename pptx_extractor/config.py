"""
Configuration management for PowerPoint Context Extractor.
"""

import json
import logging
from pathlib import Path
from typing import Dict, Any, Optional

logger = logging.getLogger(__name__)

class Config:
    """Configuration manager for the PowerPoint Context Extractor."""
    
    def __init__(self, config_path: Optional[str] = None):
        """
        Initialize configuration.
        
        Args:
            config_path: Path to configuration file. If None, uses default location.
        """
        self._config = {}
        self._config_path = self._find_config_path(config_path)
        self._load_config()
    
    def _find_config_path(self, config_path: Optional[str]) -> Path:
        """Find the configuration file path."""
        if config_path:
            return Path(config_path)
        
        # Look in several locations
        possible_paths = [
            Path("config.json"),  # Current directory
            Path(__file__).parent.parent / "config.json",  # Project root
            Path.home() / ".pptx_extractor" / "config.json",  # User home
        ]
        
        for path in possible_paths:
            if path.exists():
                return path
        
        # Default to project root
        return Path(__file__).parent.parent / "config.json"
    
    def _load_config(self):
        """Load configuration from file."""
        try:
            if self._config_path.exists():
                with open(self._config_path, 'r', encoding='utf-8') as f:
                    self._config = json.load(f)
                logger.info(f"Loaded configuration from {self._config_path}")
            else:
                logger.warning(f"Configuration file not found at {self._config_path}, using defaults")
                self._config = self._get_default_config()
        except Exception as e:
            logger.error(f"Error loading configuration: {e}")
            self._config = self._get_default_config()
    
    def _get_default_config(self) -> Dict[str, Any]:
        """Get default configuration when file is not available."""
        return {
            "cli_defaults": {
                "output_directory": "./output",
                "image_format": "png",
                "dpi": 300,
                "verbose": False,
                "recommend": False,
                "recommendation_method": "text",
                "llm_provider": "anthropic"
            },
            "timeouts": {
                "base_timeout_seconds": 60,
                "per_slide_basic_seconds": 5,
                "per_slide_recommendation_seconds": 20,
                "max_timeout_seconds": 600
            },
            "api_settings": {
                "anthropic": {
                    "model": "claude-3-haiku-20240307",
                    "max_tokens": 200,
                    "temperature": 0.7
                },
                "google": {
                    "model": "gemini-2.5-pro-preview-06-05"
                }
            }
        }
    
    def get(self, key_path: str, default: Any = None) -> Any:
        """
        Get configuration value using dot notation.
        
        Args:
            key_path: Dot-separated path to the configuration value (e.g., 'api_settings.anthropic.model')
            default: Default value if key is not found
            
        Returns:
            Configuration value or default
        """
        keys = key_path.split('.')
        value = self._config
        
        try:
            for key in keys:
                value = value[key]
            return value
        except (KeyError, TypeError):
            return default
    
    def set(self, key_path: str, value: Any):
        """
        Set configuration value using dot notation.
        
        Args:
            key_path: Dot-separated path to the configuration value
            value: Value to set
        """
        keys = key_path.split('.')
        config = self._config
        
        # Navigate to the parent of the target key
        for key in keys[:-1]:
            if key not in config:
                config[key] = {}
            config = config[key]
        
        # Set the final value
        config[keys[-1]] = value
    
    def save(self, config_path: Optional[str] = None):
        """
        Save configuration to file.
        
        Args:
            config_path: Path to save configuration. If None, uses current config path.
        """
        save_path = Path(config_path) if config_path else self._config_path
        
        try:
            # Ensure directory exists
            save_path.parent.mkdir(parents=True, exist_ok=True)
            
            with open(save_path, 'w', encoding='utf-8') as f:
                json.dump(self._config, f, indent=2)
            
            logger.info(f"Configuration saved to {save_path}")
        except Exception as e:
            logger.error(f"Error saving configuration: {e}")
    
    def get_cli_defaults(self) -> Dict[str, Any]:
        """Get CLI argument defaults."""
        return self.get('cli_defaults', {})
    
    def get_timeout_config(self) -> Dict[str, int]:
        """Get timeout configuration."""
        return self.get('timeouts', {})
    
    def get_api_config(self, provider: str) -> Dict[str, Any]:
        """Get API configuration for a specific provider."""
        return self.get(f'api_settings.{provider}', {})
    
    def get_supported_formats(self) -> Dict[str, list]:
        """Get supported file formats."""
        return self.get('supported_formats', {})
    
    def calculate_timeout(self, num_slides: int, has_recommendations: bool = False) -> int:
        """
        Calculate timeout based on configuration.
        
        Args:
            num_slides: Number of slides to process
            has_recommendations: Whether AI recommendations will be generated
            
        Returns:
            Timeout in seconds
        """
        timeout_config = self.get_timeout_config()
        
        base_timeout = timeout_config.get('base_timeout_seconds', 60)
        per_slide_basic = timeout_config.get('per_slide_basic_seconds', 5)
        per_slide_rec = timeout_config.get('per_slide_recommendation_seconds', 20)
        max_timeout = timeout_config.get('max_timeout_seconds', 600)
        
        per_slide_time = per_slide_basic
        if has_recommendations:
            per_slide_time += per_slide_rec
        
        total_timeout = base_timeout + (num_slides * per_slide_time)
        return min(total_timeout, max_timeout)


# Global configuration instance
_config_instance = None

def get_config(config_path: Optional[str] = None) -> Config:
    """
    Get the global configuration instance.
    
    Args:
        config_path: Path to configuration file (only used on first call)
        
    Returns:
        Configuration instance
    """
    global _config_instance
    
    if _config_instance is None:
        _config_instance = Config(config_path)
    
    return _config_instance

def reload_config(config_path: Optional[str] = None):
    """
    Reload the global configuration.
    
    Args:
        config_path: Path to configuration file
    """
    global _config_instance
    _config_instance = Config(config_path)