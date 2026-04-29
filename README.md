# DiscereNow – Support & Usage Guide

## Overview

DiscereNow is a **Word-to-SCORM content pipeline** designed for instructional designers and learning professionals.

It enables you to:

1. Structure learning content directly inside **Microsoft Word** using a standardized format (via Add-in)
2. Process and transform that structure using **DiscereNow Studio**
3. Export the final course as:
   - **SCORM package** (for LMS)
   - **Web version** (for hosting platforms like Vercel)

---

## Part 1 — Word Add-in (Content Structuring)

### What the Add-in Does

The DiscereNow Word Add-in allows you to **build structured learning content using predefined blocks**, ensuring consistency and compatibility with the Studio parser.

Instead of writing free-form documents, you create **semantic blocks** that represent real learning components.

---

### Available Content Blocks

#### 1. Paragraph
- Standard text content
- Used for explanations, instructions, and narrative

#### 2. Image + Text
- Two-column layout (text + image)
- Supports optional image (use marker when not present)

#### 3. Accordion
- Expandable sections
- Ideal for chunking content

#### 4. Tabs
- Content divided into tabbed navigation
- Good for comparisons or grouped topics

#### 5. Quiz
- Question + answers
- Supports:
  - Single choice
  - Multiple choice

⚠️ **Limitation:**  
Quiz blocks currently **do NOT generate score tracking in SCORM**.

#### 6. Video
- Embed video content

#### 7. Callout
- Highlighted message (tips, warnings, notes)

#### 8. Button
- Navigation or interaction trigger (e.g., “Continue”)

#### 9. Cards
- Grouped content blocks in card format

#### 10. Flip Cards
- Front/back interaction

---

### How to Use the Add-in

1. Open Microsoft Word
2. Launch the **DiscereNow Add-in**
3. Insert blocks using the interface
4. Fill in content inside each structured block
5. Follow the defined formatting rules (important for parsing)
6. Save the `.docx` file

---

## Part 2 — DiscereNow Studio (Processing & Export)

⚠️ **Status: Beta**

DiscereNow Studio is currently in **beta** and will receive continuous updates with new features and improvements.

Temporary download link:  
👉 https://placeholder-link-for-studio-download.com

---

### What the Studio Does

The Studio is responsible for:

- Reading the structured `.docx`
- Parsing all blocks
- Converting them into a course structure
- Allowing visual customization
- Exporting the final output

---

## Step-by-Step: Using DiscereNow Studio

### Step 1 — Information

You define the course metadata:

- Course name
- Short description
- Introduction
- Keywords
- Cover image
- SCORM version (e.g., 1.2)
- Completion mode

You also upload your **Word document (.docx)**.

---

### Step 2 — Review & Theming

This is where you configure the **visual identity of the course**.

#### Global Theme
- Choose a **course theme**
- Affects the **header and overall visual style**

#### Block-Level Themes

You can:

- Apply a theme **to all blocks**
- Or customize **each block individually**

---

### Step 3 — Export

You can export the course in two formats:

#### 1. SCORM Package
- Compatible with LMS platforms
- Standard SCORM 1.2 output

#### 2. Web Version
- Static web build
- Can be deployed on platforms like:
  - Vercel
  - Netlify
  - Any static server

---

## Current Limitations

- Quiz blocks do **not support score tracking in SCORM yet**
- Studio is under active development (beta phase)

---

## Recommended Workflow

1. Structure content in Word using the Add-in
2. Validate block usage and formatting
3. Import into DiscereNow Studio
4. Apply themes and review structure
5. Export as:
   - SCORM → LMS deployment
   - Web → External hosting

---

## Final Notes

DiscereNow bridges content creation and deployment, allowing instructional designers to work in Word while producing structured, scalable digital learning experiences.
