from src.email_reader import fetch_emails
from src.summary import create_summary

def main():
    fetch_emails()
    create_summary()

if __name__ == "__main__":
    main()