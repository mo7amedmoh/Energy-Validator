import ExcelJS from 'exceljs';

/**
 * Extracts all images from a workbook file.
 * @param {File} file The original Excel File object from input
 * @returns {Promise<Array>} List of images with their metadata and base64 data
 */
export async function extractImagesFromExcel(file) {
  const workbook = new ExcelJS.Workbook();
  const buffer = await file.arrayBuffer();
  await workbook.xlsx.load(buffer);
  
  const images = [];
  
  workbook.eachSheet((worksheet, sheetId) => {
    // Note: ExcelJS access to images is via worksheet.getImages()
    const sheetImages = worksheet.getImages();
    
    sheetImages.forEach((img) => {
      const imageNode = workbook.getImage(img.imageId);
      
      images.push({
        sheetName: worksheet.name,
        type: imageNode.extension, // e.g. 'png', 'jpeg'
        buffer: imageNode.buffer,
        base64: `data:image/${imageNode.extension};base64,${imageNode.buffer.toString('base64')}`,
        range: {
          tl: img.range.tl, // Top Left
          br: img.range.br  // Bottom Right
        }
      });
    });
  });
  
  return images;
}
