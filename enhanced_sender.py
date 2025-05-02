import requests
import json
import time
import datetime
from typing import Optional, Dict, Any, List
import logging

logger = logging.getLogger(__name__)

class EnhancedWAHAClient:
    def __init__(self, base_url: str = "http://23.23.209.128", session: str = "us-phone-bot"):
        self.base_url = base_url
        self.session = session

    def _make_request(self, endpoint: str, payload: Dict[str, Any]) -> Optional[Dict]:
        """Make a request to WAHA API"""
        url = f"{self.base_url}/api/{endpoint}"
        try:
            response = requests.post(url, json=payload)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            logger.error(f"Error in API request to {endpoint}: {e}")
            return None

    def _format_chat_id(self, chat_id: str) -> str:
        if not (chat_id.endswith('@c.us') or chat_id.endswith('@g.us')):
            # If it contains @g.us, it's a group
            if '@g.us' in chat_id:
                return chat_id
            # Otherwise, it's a personal chat
            return f"{chat_id}@c.us"
        return chat_id

    def send_message_with_typing(self, chat_id: str, message: str, typing_time: int = 2) -> Optional[Dict]:
        """Send a message with typing indicator"""
        chat_id = self._format_chat_id(chat_id)
        
        try:
            # Start typing
            self.start_typing(chat_id)
            
            # Wait for specified typing time
            time.sleep(typing_time)
            
            # Stop typing
            self.stop_typing(chat_id)
            
            # Send message
            result = self.send_message(chat_id, message)
            
            return result
            
        except Exception as e:
            logger.error(f"Error in send_message_with_typing: {e}")
            return None

    def send_message(self, chat_id: str, message: str) -> Optional[Dict]:
        """Send a message"""
        payload = {
            "chatId": self._format_chat_id(chat_id),
            "text": message,
            "session": self.session
        }
        return self._make_request("sendText", payload)

    def start_typing(self, chat_id: str) -> Optional[Dict]:
        """Start typing indicator"""
        payload = {
            "chatId": self._format_chat_id(chat_id),
            "session": self.session
        }
        return self._make_request("startTyping", payload)

    def stop_typing(self, chat_id: str) -> Optional[Dict]:
        """Stop typing indicator"""
        payload = {
            "chatId": self._format_chat_id(chat_id),
            "session": self.session
        }
        return self._make_request("stopTyping", payload)

    def send_stakeholder_report(self, group_id: str, report_data: List[Dict[str, Any]], date_str: Optional[str] = None) -> Optional[Dict]:
        """
        Send a formatted stakeholder report to a WhatsApp group
        
        Args:
            group_id: WhatsApp group ID (e.g., '120363418230720597@g.us')
            report_data: List of dictionaries containing stakeholder report data
            date_str: Optional date string for the report (defaults to today's date)
            
        Returns:
            API response as a dictionary
        """
        if not date_str:
            date_str = datetime.date.today().strftime("%d-%b-%Y")
        
        # Format the report message
        message = self._format_stakeholder_report(report_data, date_str)
        
        # Send the message to the group
        logger.info(f"Sending stakeholder report to group {group_id}")
        return self.send_message_with_typing(group_id, message, typing_time=3)
    
    def _format_stakeholder_report(self, report_data: List[Dict[str, Any]], date_str: str) -> str:
        """
        Format stakeholder report data into a readable message
        
        Args:
            report_data: List of dictionaries containing stakeholder report data
            date_str: Date string for the report
            
        Returns:
            Formatted message string
        """
        message_lines = [
            f"--- Stakeholder Report for Assignments on {date_str} ---",
            ""
        ]
        
        for stakeholder in report_data:
            name = stakeholder.get("name", "Unknown")
            total = stakeholder.get("total", 0)
            fresh = stakeholder.get("fresh", 0)
            abandoned = stakeholder.get("abandoned", 0)
            invalid_fake = stakeholder.get("invalid_fake", 0)
            cnp = stakeholder.get("cnp", 0)
            follow_up = stakeholder.get("follow_up", 0)
            ndr = stakeholder.get("ndr", 0)
            
            stakeholder_lines = [
                f"Calls assigned {name}",
                f"- Total Calls This Run - {total}",
                f"- Fresh - {fresh}",
                f"- Abandoned - {abandoned}",
                f"- Invalid/Fake - {invalid_fake}",
                f"- CNP - {cnp}",
                f"- Follow up - {follow_up}",
                f"- NDR - {ndr}",
                ""
            ]
            
            message_lines.extend(stakeholder_lines)
        
        message_lines.append(f"--- End of Report for {date_str} ---")
        
        return "\n".join(message_lines)


# Test the client if run directly
if __name__ == "__main__":
    # Set up logging
    logging.basicConfig(level=logging.INFO)
    
    client = EnhancedWAHAClient()
    
    # Test group ID (Giant Leap group from the plan)
    group_id = "120363418230720597@g.us"
    
    # Sample stakeholder report data
    sample_report = [
        {
            "name": "Deepasha",
            "total": 100,
            "fresh": 14,
            "abandoned": 53,
            "invalid_fake": 6,
            "cnp": 23,
            "follow_up": 1,
            "ndr": 3
        },
        {
            "name": "Khushi",
            "total": 100,
            "fresh": 14,
            "abandoned": 53,
            "invalid_fake": 5,
            "cnp": 28,
            "follow_up": 0,
            "ndr": 0
        },
        {
            "name": "Komal",
            "total": 100,
            "fresh": 13,
            "abandoned": 53,
            "invalid_fake": 9,
            "cnp": 21,
            "follow_up": 2,
            "ndr": 2
        }
    ]
    
    # Test date
    test_date = "01-May-2025"
    
    print("Testing send_stakeholder_report...")
    result = client.send_stakeholder_report(group_id, sample_report, test_date)
    print(f"Result: {result}")