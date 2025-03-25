
import * as XLSX from 'xlsx';
import { toast } from '@/hooks/use-toast';

export interface ExcelData {
  headers: string[];
  data: any[][];
  sheetNames: string[];
  activeSheet: string;
  sheets: Record<string, {
    headers: string[];
    data: any[][];
  }>;
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
    const sheets: Record<string, { headers: string[]; data: any[][] }> = {};
    
    // Process all sheets
    for (const sheetName of data.SheetNames) {
      const workSheet = data.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(workSheet, { header: 1 });
      
      if (jsonData.length > 0) {
        // Extract headers (first row) and the rest of the data
        const headers = jsonData[0] as string[] || [];
        const rows = jsonData.slice(1) as any[][] || [];
        
        sheets[sheetName] = {
          headers,
          data: rows
        };
      } else {
        // Handle empty sheets
        sheets[sheetName] = {
          headers: [],
          data: []
        };
      }
    }
    
    return {
      headers: sheets[activeSheet]?.headers || [],
      data: sheets[activeSheet]?.data || [],
      sheetNames: data.SheetNames,
      activeSheet,
      sheets
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
    
    // Add each sheet to the workbook
    for (const sheetName of data.sheetNames) {
      const sheetData = sheetName === data.activeSheet 
        ? { headers: data.headers, data: data.data } 
        : data.sheets[sheetName];
      
      if (sheetData) {
        // Create a worksheet from the data
        const worksheet = XLSX.utils.aoa_to_sheet([sheetData.headers, ...sheetData.data]);
        
        // Add the worksheet to the workbook
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
      }
    }
    
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
 * Change active sheet and load its data
 */
export const changeActiveSheet = (data: ExcelData, sheetName: string): ExcelData => {
  if (!data.sheetNames.includes(sheetName)) {
    return data;
  }
  
  // Save current sheet's data
  const currentSheetData = {
    headers: data.headers,
    data: data.data
  };
  
  const updatedSheets = {
    ...data.sheets,
    [data.activeSheet]: currentSheetData
  };
  
  // Load selected sheet's data
  const newSheetData = updatedSheets[sheetName];
  
  return {
    ...data,
    headers: newSheetData.headers,
    data: newSheetData.data,
    activeSheet: sheetName,
    sheets: updatedSheets
  };
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
  
  // Update the sheets object with the modified data
  const updatedSheets = {
    ...data.sheets,
    [data.activeSheet]: {
      headers: data.headers,
      data: newData
    }
  };
  
  return {
    ...data,
    data: newData,
    sheets: updatedSheets
  };
};

