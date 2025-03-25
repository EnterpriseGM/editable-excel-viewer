
import * as XLSX from 'xlsx';
import { toast } from '@/hooks/use-toast';

export interface ExcelData {
  headers: string[];
  data: any[][];
  sheetNames: string[];
  activeSheet: string;
}

/**
 * Parse an Excel file and return a structured object with the data
 */
export const parseExcelFile = async (file: File): Promise<ExcelData | null> => {
  try {
    const data = await readExcelFile(file);
    if (!data || !data.SheetNames || data.SheetNames.length === 0) {
      throw new Error('Invalid Excel file');
    }

    const activeSheet = data.SheetNames[0];
    const workSheet = data.Sheets[activeSheet];
    const jsonData = XLSX.utils.sheet_to_json(workSheet, { header: 1 });
    
    // Extract headers (first row) and the rest of the data
    const headers = jsonData[0] as string[];
    const rows = jsonData.slice(1) as any[][];
    
    return {
      headers,
      data: rows,
      sheetNames: data.SheetNames,
      activeSheet,
    };
  } catch (error) {
    console.error('Error parsing Excel file:', error);
    toast({
      title: 'Error',
      description: 'Failed to parse the Excel file. Please try a different file.',
      variant: 'destructive',
    });
    return null;
  }
};

/**
 * Read the Excel file and create a workbook object
 */
const readExcelFile = (file: File): Promise<XLSX.WorkBook> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        resolve(workbook);
      } catch (error) {
        reject(error);
      }
    };
    
    reader.onerror = (error) => reject(error);
    reader.readAsArrayBuffer(file);
  });
};

/**
 * Export the data back to an Excel file
 */
export const exportToExcel = (data: ExcelData, fileName: string = 'spreadsheet-export.xlsx') => {
  try {
    // Create a new workbook
    const workbook = XLSX.utils.book_new();
    
    // For the active sheet, create a worksheet from the data
    const worksheet = XLSX.utils.aoa_to_sheet([data.headers, ...data.data]);
    
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, data.activeSheet);
    
    // Write the workbook and save the file
    XLSX.writeFile(workbook, fileName);
    
    toast({
      title: 'Success',
      description: `File saved as ${fileName}`,
    });
  } catch (error) {
    console.error('Error exporting Excel file:', error);
    toast({
      title: 'Error',
      description: 'Failed to export the Excel file.',
      variant: 'destructive',
    });
  }
};

/**
 * Update the data at a specific cell
 */
export const updateCellValue = (
  data: ExcelData,
  rowIndex: number,
  columnIndex: number,
  value: any
): ExcelData => {
  const newData = [...data.data];
  
  // If the row doesn't exist, create it
  if (!newData[rowIndex]) {
    newData[rowIndex] = Array(data.headers.length).fill('');
  }
  
  // Update the value
  newData[rowIndex][columnIndex] = value;
  
  return {
    ...data,
    data: newData,
  };
};
