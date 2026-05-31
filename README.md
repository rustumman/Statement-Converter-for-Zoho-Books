# Statement Converter for Zoho Books

> **I built this. I am not a developer.**

Every month I was spending 20–30 minutes manually reformatting HDFC bank statements before importing them into Zoho Books. I described the problem to an AI, showed it my actual files, and by the end of an afternoon I had a working tool.

This is what AI-assisted building looks like for product people in 2025.

---

## What it does

Upload your HDFC statement (credit card or bank account) and download a Zoho Books-ready Excel file — with the correct columns, correct format, and a filename that tells you exactly what's inside.

- Auto-detects credit card vs. account statements
- Extracts payee names from UPI, NEFT, and bill payment narrations
- Names output files like `cc_1445_020426_020526.xlsx` so you always know which account and period
- Supports multiple file uploads at once

## Run it

```bash
pip install -r requirements.txt
streamlit run app.py
```

Or use the hosted version on [Streamlit Community Cloud](#).

---

## On building without a technical background

This repo is an honest example of what product people can and can't do with AI today.

**The limitations are real.** I can't debug things that break at a level I don't understand. I can't extend this in directions the AI didn't anticipate. It works for my specific bank and my specific format — it's not something anyone else can easily maintain without help.

**But the possibilities are also real.** For years, product people had ideas that stayed ideas — because we couldn't build, and getting engineering time meant competing with a roadmap. That gap is closing. Not disappearing. Closing.

What I built here isn't production software. But it solves a real problem, in a tool I control, on my timeline. That used to be impossible for someone like me.

The question worth asking isn't *"will AI replace developers."* It's — **what do product people build now that the small things are finally within reach?**
