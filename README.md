# PDF Merger Tool

A modern, browser-based tool to combine, split, and organize PDF files. This application runs entirely in your browser, ensuring your documents remain private and secure.

## Features

- **100% Client-Side Processing**: No files are uploaded to any server. All processing happens locally in your browser.
- **Drag & Drop Interface**: Easily upload multiple PDF files by dragging them into the drop zone.
- **Visual Page Management**:
  - View thumbnails of all pages.
  - Reorder pages with drag-and-drop.
  - **Two Drag Modes**: Move individual pages or move entire files (keeping all pages of a file together).
- **Rotation & Organization**:
  - Rotate individual pages or entire files by 90 degrees.
  - Remove specific pages or whole files.
  - Reset sorting to the original file order.
- **Merge & Download**: Combine your organized pages into a single PDF with a custom filename.
- **Privacy Focused**: Built with security in mind, ensuring no data leakage.

## How to Use

1.  **Open the Application**: Simply open `pdf_merger.html` in any modern web browser (Chrome, Edge, Firefox, Safari).
2.  **Upload PDFs**: Drag and drop PDF files into the upload area or click to browse.
3.  **Organize**:
    - Drag pages to reorder.
    - Use the toggle switch to swap between "Page" mode (move single pages) and "File" mode (move groups of pages).
    - Use the rotate icons to fix orientation.
    - Use the trash icons to delete unwanted pages or files.
4.  **Merge**: Enter a desired filename (optional) and click "Merge & Download" to save your new PDF.

## Dependencies

This tool uses the following libraries via CDN (no installation required):

-   [Tailwind CSS](https://tailwindcss.com/) for styling.
-   [Lucide Icons](https://lucide.dev/) for iconography.
-   [PDF-lib](https://pdf-lib.js.org/) for PDF modification and merging.
-   [PDF.js](https://mozilla.github.io/pdf.js/) for rendering PDF thumbnails.
-   [SortableJS](https://sortablejs.github.io/Sortable/) for drag-and-drop functionality.
-   [Google Fonts](https://fonts.google.com/) (Inter font).

## Author

**Andrin Senn**
Data Engineer, Data & Analytics Team

## License

This project is a standalone utility tool.
