from tasks.pull_email import _exec as pull_email
from tasks.schedule_gcr import _exec as schedule_gcr

print("Thank you for using Rusty, developed by Christian Sweat. This project is still in BETA and is under limited development.")
print("\tPlease read the file called Rusty.README if you have not already. It contains important information.\n\n\n")

while True:
    running = "no"
    
    while running != "run":
        running = input("Put all email files inside the folder labeled \" \\input\\ \", type RUN to continue:  ").lower()

    pull_email()
    schedule_gcr()