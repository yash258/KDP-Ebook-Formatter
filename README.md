# KDP Formatter SaaS

A production-ready, full-stack SaaS application designed to instantly standardize and format raw manuscript documents into professional, Amazon KDP-ready eBooks.

Featuring a fiercely minimal, typography-driven Black & White interface and a robust Python-based formatting engine, this tool removes the headache of manual document grooming.

## Core Features

- **Algorithmic Document Cleaning:** Automatically strips out hidden tracked changes, resolves editorial comments, sanitizes highlighted text colors, and purges manual `Tab` keystrokes.
- **Premium Styling Injection:** Parses user preferences (Fiction vs. Non-Fiction, Spacing, Typography) and applies high-end extracted editorial styles (e.g., 16pt `#2E74B5` Headings, Justified alignment, and 1.5 line spacing).
- **Structural Organization:** Intelligently identifies chapter breaks, aggressively clears excess empty line breaks, injects proper Page Breaks, and dynamically generates a clickable Table of Contents (with hidden page numbers optimal for eBook e-readers).
- **Zero-Config Deployment:** Native serverless architecture built for 1-click deployments directly to Vercel.

## Tech Stack

- **Frontend**: React.js + Vite
- **Backend API**: Python + Flask (`python-docx`)
- **Styling**: Brutalist Vanilla CSS (Strict Monochrome Design System)
- **Infrastructure**: Vercel Edge Serverless Functions

## How to Deploy

This repository is pre-configured for Vercel using `vercel.json`. You do not need to manually configure any build commands or routing.

1. Push this repository to GitHub.
2. Go to [Vercel](https://vercel.com/) and click **Add New Project**.
3. Import your GitHub repository.
4. Deploy! Vercel will automatically compile the React frontend and securely host the Python backend on its Edge network.

## Local Development

If you wish to edit or test the application locally, you will need **Node.js** and **Python 3.9+** installed on your system.

1. **Start the Frontend:**
   ```bash
   npm install
   npm run dev
   ```
2. **Start the Backend:**
   ```bash
   pip install -r requirements.txt
   python api/index.py
   ```
