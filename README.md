# AutoDeckAI – Automated PowerPoint Generator with Web + LLM Integration

## 📌 Overview
SlideSmith is a Python-based tool that automatically generates professional PowerPoint presentations from a given topic.  
It combines **Google Custom Search**, **web scraping**, and **Large Language Models (LLMs)** to:
- Collect relevant content from the web
- Extract and summarize information
- Generate a structured 7-slide outline (Title, Overview, 4 Key Sections, Takeaways, and Sources)
- Convert the outline into a styled PowerPoint deck

---

## 🚀 Features
- 🔍 **Web Search Integration**: Uses Google Custom Search API to retrieve relevant sources.  
- 📑 **Content Extraction**: Extracts clean text from web pages using BeautifulSoup.  
- 🤖 **LLM-powered Slide Generation**: Synthesizes structured slide outlines with GPT models.  
- 🎨 **Auto-Styled PPTX Creation**: Generates visually consistent slides with custom formatting.  
- 📅 **Auto Footers**: Adds date and page numbers to each slide.  

---

## 🛠️ Installation

Clone this repository and install the dependencies:

```bash
git clone https://github.com/yourusername/slidesmith.git
cd slidesmith

pip install -r requirements.txt
```

---

## ⚙️ Configuration

Set your credentials inside the script:

```python
api_key = "YOUR_GOOGLE_API_KEY"
search_engine_id = "YOUR_SEARCH_ENGINE_ID"
openai_api_key = "YOUR_OPENAI_KEY"
```

---

## 📊 Slide Structure

The generated presentation has 7 slides:
1. Title Slide – Topic title only
2. Overview – 4–6 key introduction bullets
   3–6. Key Sections – Each with a heading + 3–6 bullets
3. Takeaways – 3–5 concise summary bullets
4. Sources – Web sources used in the presentation
   

