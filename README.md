# OmniData AI

A local-first analytics MVP. Uploaded files are processed in memory and are not saved to disk.

## Run locally

1. Install Node.js 20 or newer.
2. In this folder, run `npm install`.
3. Optional: copy `.env.example` to `.env` and add an OpenAI API key. The key stays on your computer.
4. Run `npm start` and open <http://localhost:3000>.

Without an AI key, CSV/TSV/JSON/XLSX, TXT, DOCX, and text-based PDF analysis and image OCR work locally. XLSX uses its first populated worksheet. Audio transcription requires an OpenAI key and the explicit upload consent checkbox. AI-written insights are opt-in and send up to 20 sample records (or extracted text) to OpenAI only after you click **Generate AI insights** and confirm.

The app supports `.csv`, `.tsv`, `.json`, `.xlsx`, `.txt`, `.docx`, text-based `.pdf`, common raster images (`.png`, `.jpg`, `.jpeg`, `.webp`, `.tif`, `.tiff`), and audio files supported by OpenAI transcription (such as `.mp3`, `.wav`, `.m4a`, and `.webm`). Legacy `.xls`/`.doc` and scanned PDFs are not supported by this MVP.

## Notes

- Files are held in memory only for the duration of a request; the server does not persist uploads.
- The default upload size limit is 20 MB. Dataset summaries use the entire parsed file; the dashboard preview and CSV export are limited to the first 1,000 records.
- OCR runs locally using Tesseract.js and downloads the English language model the first time it is used.
- Text extraction from scanned PDFs, legacy Office files, and media transcription without an AI provider are not available.
- The AI provider is optional; never put API keys in the frontend or source control.
