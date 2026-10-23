# AI Slide Generator

An advanced, AI-powered PowerPoint presentation generator designed to build, refine, and customize slides. It extracts structured context from user prompts, uploaded documents, webpages, or YouTube videos, then outputs fully styled PowerPoint (.pptx) decks. It also supports filling custom PowerPoint templates directly, preserving their original styles and structures.

---

## Key Features

1. **Multi-Source Intelligence**: Generate slides from standard topics, YouTube videos (up to 10 minutes), webpages, or documents (PDF, DOCX, TXT).
2. **Template Mode V2**: Upload your own corporate or educational `.pptx` presentation templates. The AI planner analyzes layout structures, formats, and shapes, then maps the content directly to shape IDs without modifying font styles or slide layouts.
3. **Interactive Frontend Editor**: A React-based slides editor enabling live layout overrides, text edits, image swapping, and structure changes.
4. **Draw.io Architecture Diagram Engine**: Renders clean, styled system architecture diagrams on topic-relevant slides using the V4 diagram layout generator and Draw.io CLI.
5. **Robust Web Scraper**: Powered by Scrapling with adaptive selectors and StealthyFetcher, allowing reliable extraction from complex JS-heavy webpages.

---

## Project Structure

```
E:\AI_Slide_Generator\
├── backend\                 # Python Backend (FastAPI)
│   ├── api\                 # API routes and controllers
│   ├── models\              # Pydantic schemas and layout definitions
│   ├── services\            # Core logic (PPTX building, scraping, summarizer, LLM client)
│   ├── tests\               # Unit and integration tests
│   ├── main.py              # FastAPI main entrypoint
│   └── config.py            # App configurations (Draw.io executable paths)
├── ui\                      # React Vite Frontend App
├── outputs\                 # Generated PowerPoint presentations
├── temp\                    # Temporary cache for templates, images, and diagrams
├── requirements.txt         # Backend Python dependencies
└── README.md                # Project documentation
```

---

## Setup & Running the Project

### 1. Environment Configuration
Create a `.env` file at the root directory of the project using the variables defined in `.env.example`:
```bash
OPENROUTER_API_KEY=your-api-key
UNSPLASH_ACCESS_KEY=your-unsplash-access-key
PIXABAY_API_KEY=your-pixabay-access-key
```

### 2. Backend Setup
1. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Run the FastAPI backend:
   ```bash
   cd backend
   python -m uvicorn main:app --reload
   ```
   The backend API will start at `http://127.0.0.1:8000`.

### 3. Frontend Setup
1. Navigate to the ui folder:
   ```bash
   cd ui
   ```
2. Install npm dependencies:
   ```bash
   npm install
   ```
3. Start the Vite React development server:
   ```bash
   npm run dev
   ```
   The web application will be accessible at `http://localhost:5173`.

---

## IMPORTANT NOTES

1. **For better results, please download the generated PPT**: The local preview in the editor shell is for review. Once downloaded and opened in Microsoft PowerPoint, all alignments, fonts, images, and layouts look even better.
2. **Architecture Diagram generation**: Check the architecture diagram checkbox in the input field to generate custom visual system architectures on relevant slides.
3. **Generation time vs. slide count**: Presentation generation times depend directly on the requested slide count and the complexity of the content.
4. **Interactive layout overrides**: After slides are generated, you can change individual slide layouts in the editor. These layout updates will be correctly reflected in the final downloaded PPTX.
5. **Get actual architecture diagrams**: Ensure you check the "Architecture Diagram" checkbox inside the input options fields to trigger rendering.
6. **YouTube video limit**: YouTube video summarization and slide extraction supports videos of up to 10 minutes in duration.
7. **Custom PowerPoint Templates**: Template accuracy varies depending on structural complexity (some templates achieve 100% filling accuracy, others around 70%).
8. **Zoom-in in editor shell**: Please zoom in the PowerPoint preview inside the editor shell to check slide content details.
9. **Add environment variables**: A `.env` file must be added to the project, configured accordingly with `.env.example`.
10. **Install Draw.io Desktop**: You must have Draw.io installed locally on your system to allow the backend to render architecture diagrams.
