# 🧾 Business Communication Agent | Consultants & Management

### Engineer user interface equipped AI Agents and integrate them with Gmail and Microsoft Word to automate business communication (drafting emails, writing cover letters, etc.).


## 📌 Table of Contents
- <a href="#overview">Overview</a>
- <a href="#tools-technologies">Tools & Technologies</a>
- <a href="#engineering-agents">Engineering AI Agents</a>
- <a href="#author-contact">Author & Contact</a>

---
<h2><a class="anchor" id="overview"></a>Overview</h2>

This project aims to design and develop user interface equipped AI Agents with an integration of Gmail and Microsoft Word to automate business communication (drafting emails, writing cover letters, etc.). This user-friendly and time efficient program will help professionals generate sophisticated emails and cover letter with decrement in drafting time.
- Design the entire workflow and Use CrewAI to develop Agents
- Integrate the agents with Gmail and Microsoft Word
- Develop a low design user interface (UI) using Tkinter 


<h2><a class="anchor" id="tools-technologies"></a>🛠️ Tools & Technologies</h2>

| Tool/Platform | Purpose |
| :--- | :--- |
| Google AI Studio | API Key |
| CrewAI | Develop AI Agents |
| Tkinkter | User Interface |
| ChatGPT & Copilot | Code Assistance |

<h2><a class="anchor" id="engineering-agents"></a>🤖 Engineering AI Agents</h2>

## Defining Workflow

![agent-workflow](https://github.com/gaurav-patil-git/07_Business_Communication_Agents/blob/main/visuals/agent-workflow.png)
  
## Developing Agents

## Step 1: Generate API Key

![add-project](https://github.com/gaurav-patil-git/07_Business_Communication_Agents/blob/main/visuals/add-project.png)

![generate-api-key](https://github.com/gaurav-patil-git/07_Business_Communication_Agents/blob/main/visuals/generate-api-key.png)
  
- Also downloaded `credentials.json`

### Step 2: Prepare and Test LLM
```
from crewai import LLM
llm = LLM(
    model= "gemini/gemini-2.0-flash",
    temperature= 0.1,
    api_key= GEMINI_API_KEY
)
llm.call("What is the full name of APJ Abdul Kalam")
```

### Step 3: Engineer Email Assistant Agent
```
from crewai import Agent, Task, Crew
agent= Agent(
    role= "email assistant",
    goal= "Improve emails and make them sound more professional.",
    backstory= "A highly experienced communication expert and skilled in professional email writing and editing.",
    llm= llm
)
task= Task(
    description= f""" Take the following rough email and rewrite it into a professional and polished version. Expand Abbreviations:
    '''{email}'''""",
    agent= agent,
    expected_output= "A professionally written email with proper formatting and content.",
)
crew= Crew(
    agents= [agent],
    tasks= [task],
)
result= crew.kickoff()
print(result)
```

### Step 4: Engineer Cover Letter Writer Agent
```
from crewai import LLM, Agent, Task, Crew
agent = Agent(
    role= "a professional cover letter writing specialist",
    goal= "Write or Improve professional cover letter content based on the provided instructions to government/authority/organization.",
    backstory= "You are an expert cover letter writer with a deep understanding of storytelling.",
    llm= llm
)
task= Task(
    description= f"""Write professional and compelling cover letter content based on following instructions:
    '''{instructions}'''
    """,
    agent= agent,
    expected_output= """
    Professional cover letter content within 500 word, following a compelling narrative. 
    Keep the letter straight to the point.
    Keep paragraphs concise and focused.
    Do not include subject line or addressee details.
    """
)
crew= Crew(
    agents= [agent],
    tasks= [task],
    verbose=True
)
result= crew.kickoff()
result.raw
```

## Integrating Tools
### Step 1: Connecting Email Agent to Gmail

![enable-api](https://github.com/gaurav-patil-git/07_Business_Communication_Agents/blob/main/visuals/enable-api.png)
  
- Install Libraries
```
pip install google-api-python-client google-auth-oauthlib google-auth-httplib2
```
- Used **Prompt 1a** to generate integrated and runnable code
- Used **Prompt 1b** & **Prompt 1c** for debugging

### Step 2: Generating `.docx` with Cover Letter Agent
```
from docx import Document

file_path = "C:/Users/Desktop/cover-letter-content-v1.docx"

cover_letter_content = str(result.raw)

doc = Document()
doc.add_paragraph(cover_letter_content)
doc.save(file_path)

print(f"Cover letter saved successfully at: {file_path}")
```
## Developing Interface

- Used **Prompt 2** to integrate both agents and develop a Tkinter GUI.

![interface](https://github.com/gaurav-patil-git/07_Business_Communication_Agents/blob/main/visuals/interface.png)


<h2><a class="anchor" id="author-contact"></a>📝 Author & Contact</h2>

**Gaurav Patil** (Data Analyst) 
- 🔗 [LinkedIn](https://www.linkedin.com/in/gaurav-patil-in/)

