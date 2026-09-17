# Outlook-Drafting-Automation-
**Python scripts intended to semi-automate the process of drafting daily Outlook emails.**

Truth be told I always wanted a reason to create some sort of automation script while at work, and I found a convenient reason to do it with this project. Simply put, at the end of each day I had to submit an email detailing the tasks that I am currently working on or update previous tasks to indicate their completion status. Therefore, instead of manually drafting the same email day in day out, or even using a template email which would still require some level of labour, I decided to automate most of this drafting process so that I won't have to spend time typing up or changing bits and details in my emails, or make little but noticeable mistakes like putting in the wrong date or adding the wrong task.

## The automation works as such

1. The login script runs first, opening the Outlook logon page in Edge.
2. The script fills in details like the username and password.
3. Once the account has been logged in, the outlook session state of the page is saved in a json file and the login script stops running.
4. (Manually done) An excel spreadsheet of daily tasks is updated to today's date.
5. The drafting script runs, opening the logged in Outlook page of the user.
5. The script will fill in content like the recipients of the email, the subject, the body and include content from the spreadsheet into a new email draft.
6. The script will keep running while waiting for the user to check, edit and send the email.
7. Once the user has sent the email and closes the Outlook page, the drafting script will stop.

## Setup

1. Download a code editor (VSCode is preferable but any editor of your choice will work).
2. Download [Python](https://www.python.org/downloads/), ideally the newer versions as I am running my scripts in a Python 3.13 environment.
3. (Optional) If using VSCode download the Python and Python Environment extensions.
4. (Optional) Recommended to create a Python environment to isolate the modules to stay within your automation folder.
5. Install the necessary python modules in requirements.txt by running the following command in your folder's terminal: `pip install -r requirements.txt`
6. Install the required browsers for Playwright by running the following command in the same terminal: `python -m playwright install`
7. (Optional) If you would like to test your Playwright installation you can use the code provided from the [Playwright for Python Documentation](https://playwright.dev/python/docs/intro).
8. Change the yaml file names according to your needs.

## Configuration

Two YAML files sit alongside the script. Neither is committed to this repo —
`config.yaml` contains personal details, so copy `config.example.yaml` to
`config.yaml` and fill in your own values.

### config.yaml

| Key | Description | Example |
|---|---|---|
| `email` | Your outlook email | `your_email@example.com` |
| `password` | The password to your outlook email | `Password123` |
| `user_name` | Your name, used in the email subject line | `Jane Doe` |
| `signature` | Your signature to be added at the end of the email body | `Work Signature` |
| `excel_path` | Full path to the task tracker workbook | `C:\Users\you\Documents\TaskTracker.xlsx` |
| `sheet_name` | The name of the excel sheets in your task spreadsheet (refer to the excel spreadsheet example of how it should look like) | ` Tasks` `Selected_Reports` `Recipients` |

### template.yaml

Holds the email wording, kept separate from config so the text can be
reworded without touching any paths or personal details.

| Key | Description | Available placeholders |
|---|---|---|
| `subject` | Email subject line | `{user_name}`, `{date_long}` |
| `body` | Greeting text above the task table | `{day_name}`, `{date_short}` |

Placeholders are filled in at runtime, e.g. `{date_long}` becomes
`15 September 2026`.

## Scheduling

In order to make the scripts automated (or at least not require you running them manually), scheduling them to run at your desired specified time is necessary. The setup below is for Windows users only.

Note, each script has to be scheduled individually, so if you would like to run both scripts in an automated fashion then you will need to schedule two tasks. The login script should be scheduled before the email script as mentioned earlier.

1. Open Task Scheduler either using the run command or Windows search.
2. Select the `Create Task` button in the `Actions` tab (do not select "Create Basic Task").
3. Provide a name and description.
4. Go to the `Triggers` tab and select `New...`:
    - Leave the `Begin the task:` option as is.
    - Change the `Settings` so that the task is triggered `Weekly` (or choose `Daily` if you want the task to trigger everyday).
    - Tick the weekday options so that the task is triggered on those days only (unless if you want them to run on different days).
    - Change the start time to your preferred time for when the task should trigger.
    - (Optional) If you tend to forget that you have a task running after several days then tick the `Stop task if it runs longer than:` option in Advanced settings.
5. Go to the `Actions` tab and select `New...`:
    - Under the `Settings` field, in the `Program/script:` box fill in the path to the batch file which contains your specified script.
    - Refer to the comments in the batch files to configure them according to your needs.
6. Go to the `Conditions` tab and untick both `Stop if the computer swtiches to battery power` and `Start the task only if the computer is on AC power` options.
7. Click the `OK` button below once everything has been set up.