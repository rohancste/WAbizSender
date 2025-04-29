# Message templates for WhatsApp API
# Clean, reusable, and easy to extend for new scenarios

from typing import Dict

class MessageTemplates:
    # COD Confirmation template
    COD_CONFIRMATION = (
        "Hello {name},\n\n"
        "Thank you for choosing {storename} 🙌\n\n"
        "Please confirm your COD order before we proceed:\n\n"
        "Order ID: {orderID}\n"
        "Order Value: INR {amount}\n"
        "Order Items: {product_name}\n\n"
        "Wishing you a delightful shopping experience 💕"
    )

    # Promo/Discount template
    PROMO_MESSAGE = (
        "Hi {name},\n\n"
        "Get ₹{discount} off if you confirm your order now!\n"
        "Order ID: {orderID}\n"
        "Order Value: INR {amount}\n"
        "Order Items: {product_name}\n\n"
        "Hurry, offer valid for a limited time only! 🎉"
    )

    @staticmethod
    def render(template: str, data: Dict[str, str]) -> str:
        """
        Replace placeholders in the template with actual data.
        :param template: Template string with placeholders.
        :param data: Dictionary with keys matching placeholders.
        :return: Rendered message string.
        """
        try:
            return template.format(**data)
        except KeyError as e:
            raise ValueError(f"Missing placeholder in data: {e}")

# Example usage:
# msg = MessageTemplates.render(MessageTemplates.COD_CONFIRMATION, {
#     "name": "Shruti",
#     "storename": "thecaajustore",
#     "orderID": "#1855",
#     "amount": "619.05",
#     "product_name": "Caaju™ Smart Lakh Saver Challenge Money Box"
# })
# print(msg)