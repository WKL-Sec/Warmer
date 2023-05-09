# Warmer
Selenium-based Python script to automate sending emails to warm up your sender reputation and improve email deliverability

Why warm up? 
- One of the factors spam filters reject emails from new tenants is a filtering technique known as graylisting: When new senders appear, they're treated more suspiciously than senders with a previously-established history of sending email messages (think of it as a probation period). More on this [here](https://learn.microsoft.com/en-us/exchange/troubleshoot/email-delivery/ndr/fix-error-code-451-4-7-500-699-asxxx-in-exchange-online)

**Supported Email Providers**
- Outlook/Office365

**Blog Post**
- ["Can't Stop the Phish" - Tips for Warming Up Your Email Domain Right](https://whiteknightlabs.com/2023/05/09/cant-stop-the-phish-tips-for-warming-up-your-email-domain-right/)

## Requirements
To run the script, you need to install the required Python packages. You can install these packages using pip:

```bash
pip3 install -r requirements.txt
```

## Usage
To warmup an account by sending 5 emails to a single target, use the following command format:

```bash
python3 Warmer.py -u <Email ID to Warmup> -p <Password> -T <target@contoso.com> -m 3 -x 5
```
To warmup an account against a list of targets from a file, use the following command format:

```bash
python3 Warmer.py -u <Email ID to Warmup> -p <Password> -t <send_to_list.txt> -m 1
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
