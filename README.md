<img width="1573" height="763" alt="final_email_view" src="https://github.com/user-attachments/assets/ea09ccd7-6b11-47b2-9f00-44b3099231d6" /># Outlook-Drafting-Automation-
**Python scripts intended to semi-automate the process of drafting daily Outlook emails.**

Truth be told I always wanted a reason to create some sort of automation script while at work, and I found a convenient reason to do it with this project. Simply put, at the end of each day I had to submit an email detailing the tasks that I am currently working on or update previous tasks to indicate their completion status. Therefore, instead of manually drafting the same email day in day out, or even using a template email which would still require some level of labour, I decided to automate most of this drafting process so that I won't have to spend time typing up or changing bits and details in my emails, or make little but noticeable mistakes like putting in the wrong date or adding the wrong task.

## The automation works as such 

### For personal/free accounts

1. (Manually done) The excel spreadsheet of tasks is updated by the user.
2. `full_script.py` runs, opening an Outlook page in Chrome (or any Chromium-based browser installed)
3. The script first fills in the email and password to log into the user's account.
<img width="1583" height="978" alt="email_and_password" src="https://github.com/user-attachments/assets/6504012f-d2f4-4231-81ed-171ceb095330" />

4. The script will create a new mail and add in content like recipients, email subject, message body (which includes content from the spreadsheet if it is filled), and signature if desired.
<img width="1583" height="978" alt="drafting_mail" src="https://github.com/user-attachments/assets/fe21f505-0f5c-4acc-801f-fd9b97bdff60" />

5. (Optional) If the user chose the option to schedule the email then the script will schedule to send the email according to the specified date and time in the script.
<img width="1583" height="978" alt="sign_and_schedule" src="https://github.com/user-attachments/assets/58399997-0d69-4b88-a559-4c4265436510" />

6. The script will keep running waiting for the user to check and make changes to the email if necessary.
7. Once the user is satisfied and has sent the email, they should close the Outlook page and the drafting script will stop immediately.
<img width="1573" height="763" alt="final_email_view" src="https://github.com/user-attachments/assets/d7a79157-f18b-422c-aba8-4cd625c3ae48" />


### For organization accounts

1. `login_script.py` runs first, opening the Outlook logon page in Edge.
2. The script fills in details like the username and password.
3. Once the account has been logged in, the outlook session state of the page is saved in a json file and the login script stops running.
4. (Manually done) The excel spreadsheet of daily tasks is updated to today's date.
5. `drafting_script.py` runs, opening the logged in Outlook page of the user.
6. The script will fill in content like the recipients of the email, the subject, the body and include content from the spreadsheet into a new email draft.
7. (Optional) If the user chose the option to schedule the email then the script will schedule to send the email according to the specified date and time in the script.
8. The script will keep running while waiting for the user to check, edit and send the email.
9. Once the user has sent the email and closes the Outlook page, the drafting script will stop.

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
| `email` | Your outlook email | `"your_email@example.com"` |
| `password` | The password to your outlook email | `"Password123"` |
| `user_name` | Your name, used in the email subject line | `"Jane Doe"` |
| `add_signature` | If you want to add a signature at the end of the body, true to add, false to leave it empty | `true` |
| `signature` | Your signature to be added at the end of the email body | `"Work Signature"` |
| `schedule_email` | If you want to schedule when to send the email, true to schedule the email to be sent, false to send it yourself manually | `false` |
| `schedule_date` | The date of when you would like to send the email if scheduled, can either be set as today or to some other date as "month/day/year" | `today` or `08/26/2026` |
| `schedule_time` | The time of when you would like to send the email in the 12-hour time format | `12:45 PM` |
| `excel_path` | Full path to the task tracker workbook | `C:\Users\you\Documents\TaskTracker.xlsx` |
| `sheet_name` | The name of the excel sheets in your task spreadsheet (refer to the excel spreadsheet example of how it should look like) | `Tasks` `Selected_Reports` `Recipients` |

### template.yaml

Holds the email wording, kept separate from config so the text can be
reworded without touching any paths or personal details.

| Key | Description | Available placeholders |
|---|---|---|
| `subject` | Email subject line | `{user_name}`, `{date_long}` |
| `body` | Greeting text above the task table | `{day_name}`, `{date_short}` |

Placeholders are filled in at runtime, e.g. `{date_long}` becomes
`15 September 2026`.

## Spreadsheet

| Sheet Name | Description | Example | 
|---|---|---|
| Tasks Sheet | Main sheet to add and update tasks on a regular basis to be included into the email if desired by the user. New headers can be added to the right of the column just before the "Add To Email" header. | <img width="1344" height="346" alt="Tasks_sheet_view" src="https://github.com/user-attachments/assets/c377e566-1e5b-425e-908a-50d3e6f0eb31" />
 |
| Task Selection Sheet | If specific tasks are to be added into the email then the user can either choose "Add" in the "Add To Email" header of the specific task or update it by typing the number for said task into this sheet | <img width="133" height="153" alt="Report_selection_sheet_view" src="https://github.com/user-attachments/assets/fc4d55be-41ca-4349-892d-e1dfa4a57da2" />
 |
| Recipients Sheet | Allows user to specify who to send and cc the email to. Note, this sheet is a must to be filled in (at least for the "To" recipient) as the script will not allow an empty recipients field when running | <img width="442" height="142" alt="Recipients_sheet_view" src="https://github.com/user-attachments/assets/cbc10c1e-ffc1-4831-8c91-a74f15bd6671" />
 |

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
