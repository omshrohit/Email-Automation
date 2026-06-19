from datetime import datetime

with open("task_log.txt", "a") as f:
    f.write(f"Started: {datetime.now()}\n")

print("Task Ran Successfully")