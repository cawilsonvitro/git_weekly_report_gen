import git
import datetime as dt
import fpdf
from zoneinfo import ZoneInfo
import win32com.client as win32
import os

#make it work with remote origins so the code doesn't need to update stuff

path = os.getcwd()



def weekly_report(path_to_repos, end, start, other_update_path):
    """
    Generates a weekly report of git commit messages for multiple repositories within a specified date range,
    and includes additional updates from an external text file.
    Args:
        path_to_repos (list of str): List of file system paths to git repositories.
        end (datetime): The end datetime for filtering commits.
        start (datetime): The start datetime for filtering commits.
        other_update_path (str): Path to the external text file containing additional updates.
    Returns:
        str: A formatted string containing the weekly report for each repository and additional updates.
    Raises:
        FileNotFoundError: If any of the specified repository paths do not exist.
    """

    updates = []
    for path_to_repo in path_to_repos:
        if not os.path.exists(path_to_repo):
            raise FileNotFoundError(f"The specified repository path does not exist: {path_to_repo}")
        else:
            repo = git.Repo(path_to_repo)
            # repo.fetch()
            # repo.pull() 
            repo_name = repo.remotes.origin.url.split('.git')[0].split('/')[-1]
            updates.append(f"{repo_name} Weekly Report\n")
            i = 0
            for commit in repo.iter_commits():
                commit_time = commit.committed_datetime
                if commit_time > start and commit_time < end:
                    update = f"\t{i}. {commit.message.strip()}"
                    updates.append(update)
                    i += 1
    other_updates = []
    
    with open("outsidereport.txt", "+r") as f:
        for line in f.readlines():
            if line.strip():
                other_updates.append(line.strip())
                 
    if other_updates != []:
        updates.append(f"{other_updates[0]}\n")
        i = 0
        for update in other_updates[1:]:
            if update.strip():
                updates.append(f"\t{i}. {update.strip()}")
                
                i += 1
    updates_as_str = ("\n").join(updates)
    

    return updates_as_str

def create_report_and_email(path, txt, email_list=""):
    """
    Creates a report file with the given text and sends it via email to a list of recipients.
    Args:
        path (str): The directory path where the report file will be saved.
        txt (str): The content of the report to be written to the file and sent in the email body.
        email_list (list of str): List of email addresses to send the report to. If empty, no email is sent.
    Raises:
        Exception: If there is an error in file writing or sending the email.
    Note:
        The report file is named with the current date in EST timezone.
        Requires 'win32com.client' for Outlook email functionality.
    """
    time = str(dt.datetime.now(ZoneInfo('EST')).strftime("%Y-%m-%d"))
    with open(f"{path}\\{time} report.txt", "w") as f:
        f.write(txt)
    
    if email_list != "":
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To =email_list
        mail.Subject = 'weekly report'
        mail.Body = txt
        mail.Send()





repos = [
    "C:\\Users\\C376038\\Documents\\python_code\\packages\\instutil_pak\\instutil",
    "C:/Users/C376038/Documents/inst_suite/python/inst_code",
    "C:\\Users\\C376038\\Documents\\python_code\\packages\\cvd_sql"
]


           

txt = weekly_report(
    path_to_repos=repos,
    end=dt.datetime.now(ZoneInfo('EST')),
    start=dt.datetime.now(ZoneInfo('EST')) - dt.timedelta(days=11),
    other_update_path="outsidereport.txt"
)

print(txt)

create_report_and_email(path, txt, "McCamy, James W <JMCCAMY@vitro.com>; Wilson, Carl <CAWILSON@vitro.com>")



          
    

# git.Repo("C:\\Users\\C376038\\Documents\\python_code\\packages\\instutil_pak\\instutil")