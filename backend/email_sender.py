import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def send_email(selected_items, email_type, location):
    # SMTP Configuration
    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 587
    EMAIL_USER = "morganhondros@gmail.com"
    EMAIL_PASSWORD = "twws rmry phqz eopu"  # App-specific password
    RECIPIENT_EMAIL = "morgan@hondros-co.com"

    # Create message container
    msg = MIMEMultipart('alternative')
    # msg['From'] = EMAIL_USER
    msg['From'] = f"Morgan Hondros <{EMAIL_USER}>"  # Include name and email
    msg['To'] = RECIPIENT_EMAIL
    msg['Subject'] = f"Verdant Creations {location}: Low Stock Alert"

    # Build HTML content with styling
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

    # Attach HTML content to email
    msg.attach(MIMEText(html_content, 'html'))

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.ehlo()
            server.starttls()
            server.login(EMAIL_USER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_USER, RECIPIENT_EMAIL, msg.as_string())
        return {"status": "success", "message": "Email sent successfully"}
    except Exception as e:
        return {"status": "error", "message": str(e)}
