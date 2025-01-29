# PDF Maker  

This Python project provides tools to:  
1. **Convert images (JPG, PNG) to PDF.**  
2. **Convert PowerPoint files (PPT, PPTX) to PDF.**  
3. **Merge multiple PDFs into a single file.**  

---

## **Features**  
- **Image to PDF Conversion**: Convert one or multiple images into a single PDF file.  
- **PowerPoint to PDF Conversion**: Convert PowerPoint presentations (PPT/PPTX) to PDF format.  
- **PDF Merging**: Combine multiple PDF files into one.  
- **Error Handling**: Robust error handling to ensure smooth execution.  

---

## **Requirements**  
- Python 3.x  
- Libraries:  
  - `PIL` (Pillow) for image processing.  
  - `comtypes` for PowerPoint to PDF conversion.  
  - `PyPDF2` for merging PDFs.  

Install the required libraries using:  
```bash
pip install pillow comtypes PyPDF2
```

---

## **Usage**  

### 1. **Convert Images to PDF**  
```python
from pdf_maker import convert_images_to_pdf

input_files = ["image1.jpg", "image2.png"]  # Add your image paths here
output_pdf = "output.pdf"
convert_images_to_pdf(input_files, output_pdf)
```

### 2. **Convert PowerPoint to PDF**  
```python
from pdf_maker import convert_ppt_to_pdf, merge_pdfs

input_files = ["presentation1.pptx", "presentation2.ppt"]  # Add your PPT/PPTX paths here
pdf_paths = convert_ppt_to_pdf(input_files)
merge_pdfs(pdf_paths, "merged_output.pdf")
```

### 3. **Merge PDFs**  
```python
from pdf_maker import merge_pdfs

pdf_files = ["file1.pdf", "file2.pdf"]  # Add your PDF paths here
output_pdf = "merged.pdf"
merge_pdfs(pdf_files, output_pdf)
```

---

## **Example**  
```python
# Convert images to PDF
convert_images_to_pdf(["photo1.jpg", "photo2.png"], "images.pdf")

# Convert PowerPoint to PDF and merge
pdf_paths = convert_ppt_to_pdf(["slide1.pptx", "slide2.ppt"])
merge_pdfs(pdf_paths, "presentation.pdf")

# Merge existing PDFs
merge_pdfs(["doc1.pdf", "doc2.pdf"], "final.pdf")
```

---

## **License**  
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.  

---

## **Contributing**  
Contributions are welcome! Feel free to fork the repository, make improvements, and submit pull requests.  

---

## **Support**  
If you encounter any issues or have questions, please open an issue on the GitHub repository.  

---

This structure provides a clear and professional overview of your project, making it easy for others to understand and use. Let me know if you need further adjustments! 😊
