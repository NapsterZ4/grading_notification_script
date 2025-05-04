import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import argparse
import os

SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587


def send_email(sender_email, sender_password, recipient_email, subject, body):
    """
    Sends an email using Office 365.
    """

    try:
        # Configure the email message
        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = recipient_email
        msg["Subject"] = subject

        msg.attach(MIMEText(body, "plain"))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, msg.as_string())
        print(f"Email sent to {recipient_email}")
    except Exception as e:
        print(f"Error sending email to {recipient_email}: {e}")


def generate_email_body(row, columns, comments):
    """
    Generates the email body with the available fields.
    """
    body = f"Hello,\n\nHere is your updated information about the course:\n\n"
    for column in columns:
        if pd.isna(row[column]):
            body += f"- {column}: no data\n"
        else:
            body += f"- {column}: {row[column]}\n"

    if comments is not None and comments != "":
        body += f"\n{str(comments)}"
    body += "\n\nBest regards,\nJorge Zapata."
    return body


def read_file(file_path) -> pd.DataFrame:
    extension = os.path.splitext(file_path)[1].lower()

    if extension == ".csv":
        df = pd.read_csv(file_path)
    elif extension in [".xlsx", ".xls"]:
        df = pd.read_excel(file_path, engine="openpyxl")
    else:
        raise ValueError(f"Format not supported: {extension}")

    return df


def main(
        file_path,
        email_column,
        sender_email,
        sender_password,
        course_name,
        comments
):
    """
    Loads the data file and sends emails to all recipients.
    """
    df = read_file(file_path)
    columns = [col for col in df.columns if col != email_column]

    if email_column not in df.columns:
        raise ValueError(
            f"The column '{email_column}' does not exist in the file."
        )

    for index, row in df.iterrows():
        recipient_email = row[email_column]
        if pd.isna(recipient_email):
            print(f"Skipping row {index}: no email address.")
            continue

        subject = f"Grading Notification for {course_name}"
        body = generate_email_body(row, columns, comments)
        send_email(
            sender_email,
            sender_password,
            recipient_email,
            subject,
            body
        )


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Automatically send emails with available data."
    )

    parser.add_argument(
        "--file_path",
        required=True,
        help="Path to the file containing the data"
    )
    parser.add_argument(
        "--email_column",
        required=True,
        help="Name of the column containing email addresses"
    )
    parser.add_argument(
        "--sender_email",
        required=True,
        help="Office 365 email address to send messages from"
    )
    parser.add_argument(
        "--sender_password",
        required=True,
        help="Password for the email (you need to use the app password)"
    )
    parser.add_argument(
        "--course_name",
        required=True,
        help="The name of your course"
    )
    parser.add_argument(
        "--comments",
        help="If you want to add comments on the email"
    )

    args = parser.parse_args()

    main(
        args.file_path,
        args.email_column,
        args.sender_email,
        args.sender_password,
        args.course_name,
        args.comments
    )
