"""
WAbizSender - WhatsApp Business API Integration for Shopify

This application provides:
1. CRM functionality for Shopify sellers
2. Automated WhatsApp messaging using the WhatsApp Business API
"""

import os
import sys

# Add the project root to the path to ensure imports work correctly
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def main():
    """Main entry point for the application."""
    print("WAbizSender - Starting application")
    
    # Import here to avoid circular imports
    from data.distribution import distribute_and_report
    
    # Run the distribution script
    distribute_and_report()
    
    # TODO: Add WhatsApp API functionality
    
if __name__ == "__main__":
    main()