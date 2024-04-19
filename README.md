# Email Sender

This Python script is used to send emails with certificate attachments to a list of recipients.

## Prerequisites

- Python 3.x
- The following Python libraries:
  - os
  - yaml
  - logging
  - smtplib
  - tqdm
  - logging.config
  - dotenv
  - numpy
  - pandas

## Installation

1. Clone the repository.
2. Install the required Python libraries using `pip`:

```bash
   pip install -r requirements.txt
```

## Usage
Create a `.env` file in the root directory of the project and add the following variables:
```python
SMTP_SERVER=your_smtp_server
SMTP_PORT=your_smtp_port
EMAIL_ADDRESS=your_email_address
EMAIL_PASSWORD=your_email_password
```
- Create a `recipients.csv` file in the root directory of the project. The CSV file should have two columns: `Name` and `Email`.
- Run the script:
  ```python
  py send_email.py
  ```
- The script will send an email to each recipient with a certificate attachment.
- The script logs the progress and any errors to a YAML formatted log file named `email_log.yaml`.
