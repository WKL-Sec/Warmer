# Warmer
Collection of Selenium-based Python scripts to automate sending emails to warm up your sender reputation and improve email deliverability

Why warm up? 
- One of the factors spam filters reject emails from new tenants is a filtering technique known as graylisting: When new senders appear, they're treated more suspiciously than senders with a previously-established history of sending email messages (think of it as a probation period). More on this [here](https://learn.microsoft.com/en-us/exchange/troubleshoot/email-delivery/ndr/fix-error-code-451-4-7-500-699-asxxx-in-exchange-online)

**Supported Email Providers**
- ProtonMail
- Outlook

![Warmer](https://user-images.githubusercontent.com/97109724/235449168-8e4d5399-c3a3-4e14-b4b7-7b4e0517a616.png)

Blog Post
- ["Can't Stop the Phish" - Tips for Warming Up Your Email Domain Right](https://whiteknightlabs.com/)

## Requirements
To run the script, you need to install the required Python packages. You can install these packages using pip:

```bash
pip3 install -r requirements.txt
```

## Usage
To run the program, use the following command format:

```bash
python3 Warmer-OWA.py -u <Email ID to Warmup> -p <Password> -t <send_to_list.txt> -m 3
```

### Command Line Arguments

The script supports the following command line arguments:

```bash
-h, --help:  Show this help message and exit
-u: Sender Outlook Email ID
-p: Sender Outlook Email Password
-T: Single Target Email ID
-t: Multiple Targets from Wordlist
-x: No. of Emails to Send (applicable only for single targets)
-m: Email Content Mode [ 1, 2, 3] where 1 = Gibberish sentence, 2 = AI-Generated, 3 = Randomly choose from pre-defined templates
```
