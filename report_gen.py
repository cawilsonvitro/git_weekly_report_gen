import git
import datetime as dt
import pytz
from zoneinfo import ZoneInfo

path_to_repo = r"C:\Users\C376038\Documents\python_code\packages\instutil_pak\instutil"  


eastern_timezone = pytz.timezone('US/Eastern')

repo = git.Repo(path_to_repo)

end = dt.datetime.now(ZoneInfo('EST'))
start = end - dt.timedelta(days=5)
updates = []

for commit in repo.iter_commits():
    commit_time = commit.committed_datetime
    if commit_time > start and commit_time < end:
        updates.append(commit.message.strip())

        print(f"Commit: {commit.hexsha}")
        print(f"Author: {commit.author.name} <{commit.author.email}>")
        print(f"Date: {commit.committed_datetime}")
        print(f"Message: {commit.message.strip()}")
        print("-" * 40)
        
        
        
print(updates)