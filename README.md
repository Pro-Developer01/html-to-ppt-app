# HTML to PowerPoint Converter (React)

A React-based web application that converts HTML code to editable PowerPoint presentations using the dom-to-pptx library.

## Features

- Paste HTML code directly into the UI
- Convert HTML to PPTX format
- Download generated PowerPoint files
- Clean, modern React interface

## Installation

```bash
npm install
```

## Usage

### Development Mode

Start the development server:

```bash
npm run dev
```

Then open http://localhost:3000 in your browser.

### Instructions

1. Paste your HTML code into the text area
2. Click "Convert to PowerPoint"
3. Your PPTX file will be downloaded automatically

## Build for Production

```bash
npm run build
npm run preview
```

## Legacy Scripts

The original Node.js conversion scripts are still available:

- `npm run serve` - Start the old static server
- `npm run convert` - Run automated conversion with Puppeteer

## Tech Stack

- React 18
- Vite
- dom-to-pptx library
- CSS3

## Notes

- The HTML should contain a `.slide-container` element for best results
- Slide dimensions are set to 10" x 5.625" (16:9 aspect ratio)
- Complex CSS and web fonts are supported
