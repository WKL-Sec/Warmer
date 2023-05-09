from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import sys
import time
import argparse
import pyfiglet
import random
import string
from termcolor import colored, cprint

import openai

############
### Modify this with your API Key

openai.api_key = ""

############


def random_word():
    # Generates a random word of length 3 to 6 consisting of random letters
    word_length = random.randint(3, 6)
    word = ''.join(random.choice(string.ascii_lowercase) for _ in range(word_length))
    return word

def random_sentence():
    # Generates a random sentence of up to 30 words
    sentence_length = random.randint(15, 30)
    sentence = ' '.join(random_word() for _ in range(sentence_length))
    return sentence.capitalize() + '.'

def send_owa_email(owa_login, owa_pass, email_list, counter, flag, email_mode):
    url_owa = "https://outlook.office.com/mail"

    print("[+] Visiting OWA Portal", end='\r')
    driver.get(url_owa)
    time.sleep(7)

    print("\n[+] Logging into Outlook as {0}".format(owa_login))
    
    driver.find_element(By.ID, "i0116").send_keys(owa_login)
    driver.find_element(By.ID, "idSIButton9").click()
    time.sleep(3)

    driver.find_element(By.ID, "i0118").send_keys(owa_pass)
    driver.find_element(By.ID, "idSIButton9").click()
    time.sleep(3)

    driver.find_element(By.ID, "idBtn_Back").click()
    time.sleep(7)

    print("[+] Authenticated successfully", end='\r')
    
    if flag == 0:
        print("[+] Composing emails for {0}".format(email_list[0]))

    subject_list = ["Weekend plans", "Happy holidays!", "Last weeks discussion", "Plan for the trip", "Team Celebrations", "Travel Itinary for Next Month", "Lunch Date", "Whiteboard Meeting Call","A few things to celebrate this week", "Debrief Call - Project Completed", "Quote for new opportunity", "Leave Details Request"]
    body_list = ["Hey there,\n\nJust wanted to check in and see if you're up for some weekend fun? I was thinking of hitting up the farmers market on Saturday and then grabbing lunch. Let me know if you're interested.", "Hey,\n\nI just wanted to wish you a happy birthday and let you know how grateful I am to have you in my life. Hope you have a great day. Cheers!", "Hi there,\n\nIt's been a while since we caught up! I was thinking of grabbing lunch next week, would you be free on Wednesday or Thursday? Let me know!\n\nRegards,\nJen", "Hi,\n\nI saw on LinkedIn that you got a promotion, congrats! I know how hard you've been working and you deserve it. Let's grab drinks to celebrate soon. Cheers!", "Hi there,\n\nHope you're doing well. I had a quick question about that project we worked on last month. Would you mind hopping on a call sometime this week to chat about it? D.", "Hey,\n\nI know you've traveled to Paris before and I was wondering if you had any tips or recommendations for my upcoming trip. I'd love to hear your thoughts. J.", "Hey,\n\nJust wanted to wish you and your family a happy holiday season! Hope you get to enjoy some time off and relaxation.", "Hi, All,\n\nJust summarising the call today:\nProcessor is only 2kb so Mark has suggested we do CTF type challenges. D", "Hi Everyone,\n\nA few things for us to celebrate this week:\n\n Owen is celebrating his birthday on Thursday\n Jess will be celebrating his 1st work anniversary on Friday\n\nBest wishes to you both, from your friends in WKL"]
 
    i = 1
    tots_sent = counter

    while i <= counter:

        if email_mode == '1':
            email_subject = ' '.join(random_word() for _ in range(3))
            email_message = "Hi there,\n\n" + random_sentence() + "\n" + random_sentence() + random_sentence() + "\n\n" + random_sentence()
        
        elif email_mode == '2':
            email_subject = random.choice(subject_list) # Removed support for OpenAI to avoid unnecessary costs
            email_message = askGPT()

        else:
            email_subject = random.choice(subject_list)
            email_message = random.choice(body_list)
        
        try:

            driver.find_element(By.CLASS_NAME, "label-186").click()
            time.sleep(10)

            if flag == 1:
                email_to = email_list[i-1]
                print("[+] Composing emails for {0}".format(email_to))

            else:
                email_to = email_list[0]
		
            driver.switch_to.active_element.send_keys(email_to + '\n' + '\n' + '\t' + '\t' + email_subject + '\t' + email_message)
            time.sleep(6)

            driver.find_element(By.XPATH, "//button[@title='Send (Ctrl+Enter)' and contains(@class, 'ms-Button--primary')]").click()
            time.sleep(6)

            print("[+] Counter: {0}".format(i), end='\r')

        except Exception as e:

             print("[!] Error on Counter {0}: {0}".format(i, e))
             tots_sent -=1

        i += 1
        time.sleep(8)

    print("[+] {0} emails completed".format(tots_sent))
    print("[+} Logging out..", end='\r')    

    driver.find_element(By.ID, "O365_MainLink_Me").click()
    time.sleep(3)

    driver.find_element(By.ID, "mectrl_body_signOut").click()
    time.sleep(3)

    print("[+] Logged out successfully", end='\r')
    time.sleep(1)


def askGPT():

    if openai.api_key == "":
        print("[!] OpenAI API key is required.")

    prompt_text = "Write a typical corporate email without a subject in 50 words that do not use any words that may trigger spam controls. Begin with the body without any trail spaces or new lines, and sign off with two new lines, followed by the name Bob at the end"
    response = openai.Completion.create( engine="text-davinci-002", prompt=prompt_text, temperature=0.6,  max_tokens=150 )
    generated_text = response.choices[0].text
    return generated_text

# Main
prebanner = pyfiglet.figlet_format("Warmer")
VERSION = colored('VERSION: 1.0 \n\n-- Written by @Firestone65 --\n', 'red', attrs=['bold'])
'Sender reputation warmer for phishing campaigns'
banner = prebanner + "\n" + VERSION
print(banner)

parser = argparse.ArgumentParser(description= '[+] Sender reputation warmer for phishing campaigns')
parser.add_argument('-u', type=str, required=True, help='Sender Outlook Email ID')
parser.add_argument('-p', type=str, required=True, help='Sender Outlook Email Password')
parser.add_argument('-T', type=str, dest="T", required=False, help='Single Target Email ID')
parser.add_argument('-t', type=str, dest="t", required=False, help='Multiple Targets from Wordlist')
parser.add_argument('-x', type=int, required=False, help='No. of Emails to Send (applicable only for single targets)')
parser.add_argument('-m', type=str, required=True, help='Email Content Mode [ 1, 2, 3] where 1 = Gibberish sentence, 2 = AI-Generated, 3 = Randomly choose from pre-defined templates')
args = parser.parse_args()

# Distinguish between single vs. multiple recipients
flag = 0

#Email Information
email_list = list()

if args.t is not None:
    emails_txt = args.t

    with open(emails_txt, 'r') as fp:
        
        flag = 1
        email_list = [line.rstrip('\n') for line in fp.readlines()]
        send_volume = len(email_list)


elif args.T is not None:
    email_list.append(args.T)
    
    if args.x is not None:
        send_volume = args.x
    else:
        print("[!] Send volume not provided. Defaulting to 1 email")
        send_volume = 1

else:
    print("[!] Required fields: Target Email ID / Target Email Wordlist")
    sys.exit()

email_mode = args.m

outlook_login = args.u
outlook_pass = args.p

# Initialize Chrome driver
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

send_owa_email(outlook_login, outlook_pass, email_list, send_volume, flag, email_mode) 

print("\n[+] Automation complete", end='\r')
#inp = input('\n ---- Hit any key to quit')

driver.quit()
