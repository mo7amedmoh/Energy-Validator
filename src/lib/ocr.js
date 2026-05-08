import { createWorker } from 'tesseract.js';

/**
 * Perform basic OCR on an image buffer or base64 string.
 * Focused on numbers for voltage/ampere readings.
 * @param {string|Buffer} imageSource Base64 string or image buffer
 * @returns {Promise<string>} Extracted text
 */
export async function performOCR(imageSource) {
  const worker = await createWorker('eng');
  
  try {
    // Configure for digits only if we know we are looking for meter readings
    // await worker.setParameters({
    //   tessedit_char_whitelist: '0123456789.VvAa:',
    // });
    
    const { data: { text } } = await worker.recognize(imageSource);
    await worker.terminate();
    return text.trim();
  } catch (error) {
    console.error("OCR Error:", error);
    await worker.terminate();
    return "Error reading text";
  }
}
