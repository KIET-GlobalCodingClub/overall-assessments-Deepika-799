import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
import pandas as pd
import os

# ========== 1. Generate Certificate ==========
def generate_certificate(name, output_file="generate_certificates.pdf"):
    width, height = A4
    c = canvas.Canvas(output_file, pagesize=A4)

    # Add certificate template image
    template = ImageReader("certificate_template.jpg")
    c.drawImage(template, 0, 0, width=width, height=height)

    # Add participant's name (adjust position according to your template)
    c.setFont("Helvetica-Bold", 32)
    c.drawCentredString(width/2, height/2, name)

    c.save()
    return output_file


# ========== 2. Send Email ==========
def send_email(to_email, subject, body, attachment_path):
    from_email = "deepikakunkatla4@gmail.com"
    from_password = "fcns ltnz ymhp mhmi"   # Use Gmail App Password

    # Email setup
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Attach certificate
    with open(attachment_path, "rb") as f:
        mime = MIMEBase('application', 'octet-stream')
        mime.set_payload(f.read())
        encoders.encode_base64(mime)
        mime.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
        msg.attach(mime)
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(from_email, from_password)
            server.send_message(msg)
        print(f"✅ Certificate sent to {to_email}")
    except Exception as e:
        print(f"❌ Failed to send to {to_email}: {e}")


# ========== 3. Main ==========
if __name__ == "__main__":
    # Load participants from Excel
    try:
        df = pd.read_excel("students.xlsx")
        
        for _, row in df.iterrows():
            Name = row["name"]
            Email = row["mail"]

            # Generate certificate
            cert_file = f"cert_{Name.replace(' ', '_')}.pdf"
            generate_certificate(Name, cert_file)

            # Send certificate via email
            send_email(
                Email,
                "Your Certificate of Completion",
                f"Hello {Name},\n\nCongratulations! Please find your certificate attached.\n\nBest Regards,\nEvent Team",
                cert_file
            )
    except FileNotFoundError:
        print("❌ Error: students.xlsx file not found. Please make sure the file exists.")
    except KeyError as e:
        print(f"❌ Column error: {e}")
        print("Available columns:", df.columns.tolist() if 'df' in locals() else "File not loaded")
    except Exception as e:
        print(f"❌ An error occurred: {e}")
