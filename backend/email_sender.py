from email.message import EmailMessage
from email.utils import formataddr
import smtplib
from flask import jsonify
from smtplib import SMTPAuthenticationError, SMTPException
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

def send_email(selected_items, email_type, location):
    msg = EmailMessage()
    msg['Subject'] = f"LOW STOCK ALERT - From Verdant Creations in {location}."
    msg['From'] = formataddr(("Morgan Hondros", "morganhondros@gmail.com"))
    msg['To'] = formataddr(("Morgan Hondros", "morgan@hondros-co.com"))

    html_content = f"""
    <html>
    <head>
        <style>
            body {{ font-family: 'Helvetica', 'Arial', sans-serif; line-height: 1.6; color: #333; }}
            .container {{ width: 100%; max-width: 600px; margin: 0 auto; padding: 20px; }}
            .header {{ background-color: #3d6428; color: white; padding: 20px; text-align: center; }}
            .content {{ background-color: #f8f9fa; padding: 20px; }}
            .list-unstyled {{ padding-left: 0; list-style-type: none; }}
            .list-item {{ border-radius: 4px; padding: 10px; margin-bottom: 10px; }}
            .vendor-name {{ font-weight: bold; color: #3d6428; margin-bottom: 10px; }}
            .product-name {{ font-weight: bold; }}
            .product-info {{ color: #6c757d; }}
            .bg-danger {{ background-color: #f8d7da; border: 1px solid #f5c6cb; }}
            .bg-warning {{ background-color: #fff3cd; border: 1px solid #ffeeba; }}
            .bg-light {{ background-color: #f8f9fa; border: 1px solid #dee2e6; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h2 style="margin: 0;">Verdant Creations {location}: Low Stock Alert</h2>
            </div>
            <div class="content">
                <p>We have a list of your products that are currently running low. Expect an order from us soon.</p>
                <ul class="list-unstyled">
    """

    for vendor, items in selected_items.items():
        html_content += f'<li class="list-item bg-light"><div class="vendor-name">{vendor}</div><ul class="list-unstyled">'
        for item in items:
            if item['daysUntilSoldOut'] <= 3:
                item_class = "bg-danger"
            elif item['daysUntilSoldOut'] <= 7:
                item_class = "bg-warning"
            else:
                item_class = "bg-light"

            html_content += f"""
                <li class="list-item {item_class}" style="padding-bottom: 10px;">
                    <span class="product-name">{item['name']}</span>
                    <div class="product-info">
                        Days until sold out: {item['daysUntilSoldOut']}
                        <br>
                        Quantity left: {item.get('remainingQty', 'N/A')}
                    </div>
                </li>
            """
        html_content += "</ul></li>"

    html_content += """
                </ul>
            </div>
        </div>
    </body>
    </html>
    """

    msg.set_content(html_content, subtype='html', charset='utf-8')

    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.set_debuglevel(1)  # Enable debug output
            server.starttls()
            server.login("morganhondros@gmail.com", "xuxn lqvi qdjo soou")
            server.send_message(msg)
        return {'message': 'Email sent successfully'}
    except SMTPAuthenticationError as e:
        return {'error': f"Authentication error: {str(e)}"}
    except SMTPException as e:
        return {'error': f"SMTP error: {str(e)}"}
    except Exception as e:
        return {'error': f"General error: {str(e)}"}
