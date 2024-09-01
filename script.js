   document.getElementById('extractButton').addEventListener('click', extractComments);

    async function extractComments() {
        const fileInput = document.getElementById('fileInput').files;
        if (fileInput.length === 0) {
            alert('Please upload at least one PDF file.');
            return;
        }

        let comments = [];
        for (const file of fileInput) {
            const pdfData = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: pdfData }).promise;

            for (let i = 0; i < pdf.numPages; i++) {
                const page = await pdf.getPage(i + 1);
                const annotations = await page.getAnnotations();

                annotations.forEach(annotation => {
                    // Filter out popups to avoid duplication
                    if (annotation.subtype !== 'Popup' && annotation.rect) {
                        comments.push({
                            FileName: file.name,
                            Page: i + 1,
                            Type: annotation.subtype,
                            Content: annotation.contents || 'No content',
                            Author: annotation.title || 'Unknown',
                            Rect: annotation.rect.join(', ') // Joining array elements for better readability
                        });
                    }
                });
            }
        }

        exportToExcel(comments);
    }

    function exportToExcel(comments) {
        // Define the worksheet and workbook
        const worksheet = XLSX.utils.json_to_sheet(comments);
        const workbook = XLSX.utils.book_new();

        // Append the worksheet to the workbook
        XLSX.utils.book_append_sheet(workbook, worksheet, "Comments");

        // Create the Excel file and trigger a download
        XLSX.writeFile(workbook, "comments.xlsx");
    }
