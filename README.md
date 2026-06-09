# WhatsappLogger
Automated issue logging from WhatsApp message to log tracking list

pip install openai openpyxl pydantic
pip install flask flask-cors

$env:DEEPSEEK_API_KEY="YOUR_API_KEY"
python app.py

Python Backend Setup
Step 1: Prepare Your Workspace
Create a new folder on your desktop (e.g., IssueLog).

Inside that folder, create a blank Excel file and name it exactly IssueLog.xlsx. (The script will look for this specific file in the same folder).

Open a terminal (PowerShell or Command Prompt) and navigate to that new folder:

PowerShell
cd C:\IssueLog
Step 2: Install Python Dependencies
The script relies on three lightweight libraries. Install them via pip by running this command in your terminal:

PowerShell
pip install openai openpyxl pydantic
openai: Communicates with the GPT-4o-mini model.
openpyxl: Reads and writes to your Excel file locally without needing Excel to be open.
pydantic: Forces the AI to output perfect data matching your columns.

Step 3: Set Your OpenAI API Key
The agent needs permission to use OpenAI's API.
Go to platform.openai.com and generate a new secret API key.
In your open PowerShell window, temporarily set the key as an environment variable so the script can authenticate:

PowerShell
$env:OPENAI_API_KEY="sk-proj-your-actual-api-key-here"
(Note: If you are using standard Command Prompt, use set OPENAI_API_KEY=sk-proj-... instead).

Step 4: Save the Python Script
Save deepseek_agent.py and app.py in the folder

Step 5: Execute the Test
With your Excel file and Python script in the same folder, and your API key set, run the script from your terminal:

PowerShell
python agent_test.py

Step 6: Verify the Output
