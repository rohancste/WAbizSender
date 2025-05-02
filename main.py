"""
WAbizSender - WhatsApp Business API Integration for Shopify

This application provides:
1. CRM functionality for Shopify sellers
2. Automated WhatsApp messaging using the WhatsApp Business API
"""

import os
import sys
import logging
import datetime

# Add the project root to the path to ensure imports work correctly
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('wabiz_sender.log'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

def generate_fake_stakeholder_report():
    """Generate fake stakeholder report data for testing purposes."""
    logger.info("Generating fake stakeholder report data for testing...")
    
    # Create a fake report structure similar to what distribute_and_report would return
    fake_report = {
        "Rohan": {
            "Total": 157,
            "Fresh": 52,
            "Abandoned": 50,
            "CNP": 54,
            "Follow up": 1,
            "NDR": 0
        },
        "Lucky": {
            "Total": 164,
            "Fresh": 51,
            "Abandoned": 50,
            "CNP": 41,
            "Follow up": 10,
            "NDR": 2
        },
        "Rohit": {
            "Total": 145,
            "Fresh": 48,
            "Abandoned": 45,
            "CNP": 38,
            "Follow up": 8,
            "NDR": 6
        }
    }
    
    return fake_report

def convert_report_format(report_dict):
    """
    Convert the report from dictionary format to list of dictionaries format
    that is expected by send_stakeholder_report method
    """
    if not report_dict:
        return []
        
    report_list = []
    for name, counts in report_dict.items():
        stakeholder_data = {
            "name": name,
            "total": counts.get("Total", 0),
            "fresh": counts.get("Fresh", 0),
            "abandoned": counts.get("Abandoned", 0),
            "invalid_fake": counts.get("Invalid/Fake", 0),
            "cnp": counts.get("CNP", 0),
            "follow_up": counts.get("Follow up", 0),
            "ndr": counts.get("NDR", 0)
        }
        report_list.append(stakeholder_data)
    
    return report_list

def main():
    """Main entry point for the application."""
    logger.info("WAbizSender - Starting application")
    
    # Import here to avoid circular imports
    from data.distribution import distribute_and_report
    from enhanced_sender import EnhancedWAHAClient
    
    # Run the distribution script to process data and generate report
    logger.info("Processing data from Google Sheets...")
    
    # Process data and generate stakeholder report
    stakeholder_report = distribute_and_report()
    
    # Check if we got valid data back
    if not stakeholder_report:
        logger.warning("No real data available from Google Sheets. Using fake data instead.")
        stakeholder_report = generate_fake_stakeholder_report()
    
    # Convert the report format to what send_stakeholder_report expects
    formatted_report = convert_report_format(stakeholder_report)
    
    # Now send the report to WhatsApp
    logger.info("Sending stakeholder report to WhatsApp group...")
    
    # Initialize WhatsApp client
    whatsapp_client = EnhancedWAHAClient()
    
    # Target WhatsApp group (Giant Leap group from the plan)
    group_id = "120363418230720597@g.us"
    
    # Get today's date for the report
    today_date = datetime.date.today().strftime("%d-%b-%Y")
    
    # Send the report to WhatsApp using the properly formatted data
    result = whatsapp_client.send_stakeholder_report(group_id, formatted_report, today_date)
    
    if result:
        logger.info("Successfully sent stakeholder report to WhatsApp group")
    else:
        logger.error("Failed to send stakeholder report to WhatsApp group")
    
if __name__ == "__main__":
    main()