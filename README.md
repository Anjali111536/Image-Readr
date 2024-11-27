// Set your API Key
const scriptProperties = PropertiesService.getScriptProperties();
const VISION_API_KEY = scriptProperties.getProperty("API_KEY");

function extractTextFromImages() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();

    // Process each row (skip header row)
    for (let row = 2; row <= lastRow; row++) { // Row 1 is the header
        const imageUrl = sheet.getRange(row, 15).getValue(); // Column O (Image URL)
        const extractedText = sheet.getRange(row, 19).getValue(); // Column S (Extracted Text)

        // Skip rows that already have extracted text
        if (imageUrl && !extractedText) {
            const processedImageUrl = convertGoogleDriveUrl(imageUrl);
            const formattedText = callVisionAPI(processedImageUrl);

            if (formattedText) {
                // Save extracted and formatted text
                sheet.getRange(row, 19).setValue(formattedText); // Save in Column S
            }
        }
    }
}

// Google Drive URL to Vision API compatible URL
function convertGoogleDriveUrl(url) {
    const fileIdMatch = url.match(/\/d\/(.*?)\//);
    if (fileIdMatch && fileIdMatch[1]) {
        return `https://drive.google.com/uc?export=download&id=${fileIdMatch[1]}`;
    }
    return url; // Return original URL if not in Google Drive format
}

function callVisionAPI(imageUrl) {
    const visionEndpoint = `https://vision.googleapis.com/v1/images:annotate?key=${VISION_API_KEY}`;

    try {
        const response = UrlFetchApp.fetch(imageUrl);
        const imageBlob = response.getBlob();
        const base64Image = Utilities.base64Encode(imageBlob.getBytes());

        // Vision API request payload
        const requestPayload = {
            requests: [
                {
                    image: {
                        content: base64Image
                    },
                    features: [
                        {
                            type: "DOCUMENT_TEXT_DETECTION"
                        }
                    ]
                }
            ]
        };

        const options = {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(requestPayload)
        };

        const visionResponse = UrlFetchApp.fetch(visionEndpoint, options);
        const jsonResponse = JSON.parse(visionResponse.getContentText());

        // Extract the full text preserving format
        if (jsonResponse.responses && jsonResponse.responses[0].fullTextAnnotation) {
            const text = jsonResponse.responses[0].fullTextAnnotation.text;
            return text; // Preserved line breaks and layout
        } else {
            return "No text found";
        }
    } catch (error) {
        Logger.log("Error processing image: " + error.message);
        return "Error processing image";
    }
}
