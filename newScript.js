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

        // Extract banner name, week, and PRF number from the filename
        const fileName = file.name;
        const bannerName = fileName.split('_')[0] + '_' + fileName.split('_')[1];
        const week = fileName.match(/WK(\d+)/)[1];
        const prfNumber = fileName.match(/PRF(\d+)/)[1];

        for (let i = 0; i < pdf.numPages; i++) {
            const page = await pdf.getPage(i + 1);
            const annotations = await page.getAnnotations();

            annotations.forEach(annotation => {
                if (annotation.subtype !== 'Popup' && annotation.rect) {
                    // Determine error type based on the comment's content
                    let errorType = 'Product_Description'; // Default category
                    let content = annotation.contents || 'No content';

                    if (content.toLowerCase().includes('price')) {
                        errorType = 'Price_Point';
                    } else if (content.toLowerCase().includes('alignment')) {
                        errorType = 'Overall_Layout';
                    } else if (content.toLowerCase().includes('image')) {
                        errorType = 'Image_Usage';
                    }

                    comments.push({
                        BannerName: bannerName, // Add banner name
                        Week: week, // Add week number
                        PRFNumber: prfNumber, // Add PRF number
                        FileName: file.name,
                        Page: i + 1,
                        Type: annotation.subtype,
                        Content: content,
                        Author: annotation.title || 'Unknown',
                        Rect: annotation.rect.join(', '),
                        ErrorType: errorType
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

document.getElementById('fileInput').addEventListener('change', function(event) {
    var fileCount = event.target.files.length;
    var fileCountText = fileCount > 0 ? fileCount + ' file(s) chosen' : 'No files chosen';
    document.getElementById('fileCount').textContent = fileCountText;
});
