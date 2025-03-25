
import axios from 'axios';

const API_URL = import.meta.env.VITE_API_URL || 'http://localhost:3001/api';

const api = axios.create({
  baseURL: API_URL,
  headers: {
    'Content-Type': 'application/json',
  },
});

export interface ApiExcelData {
  fileId: string;
  fileName: string;
  headers: string[];
  data: any[][];
  sheetNames: string[];
  activeSheet: string;
  sheets: Record<string, {
    headers: string[];
    data: any[][];
  }>;
}

export const uploadExcelFile = async (file: File): Promise<ApiExcelData> => {
  const formData = new FormData();
  formData.append('file', file);
  
  try {
    const response = await axios.post(`${API_URL}/upload`, formData, {
      headers: {
        'Content-Type': 'multipart/form-data',
      },
    });
    
    return response.data;
  } catch (error) {
    console.error('Error uploading file:', error);
    throw error;
  }
};

export const getSheetData = async (fileId: string): Promise<ApiExcelData> => {
  try {
    const response = await api.get(`/sheets/${fileId}`);
    return response.data;
  } catch (error) {
    console.error('Error getting sheet data:', error);
    throw error;
  }
};

export const changeActiveSheet = async (fileId: string, sheetName: string): Promise<ApiExcelData> => {
  try {
    const response = await api.post(`/sheets/${fileId}/change-sheet`, { sheetName });
    return response.data;
  } catch (error) {
    console.error('Error changing active sheet:', error);
    throw error;
  }
};

export const updateCellValue = async (
  fileId: string,
  rowIndex: number,
  columnIndex: number,
  value: any
): Promise<ApiExcelData> => {
  try {
    const response = await api.post(`/sheets/${fileId}/update-cell`, {
      rowIndex,
      columnIndex,
      value,
    });
    return response.data;
  } catch (error) {
    console.error('Error updating cell value:', error);
    throw error;
  }
};

export const exportExcelFile = (fileId: string): void => {
  window.open(`${API_URL}/sheets/${fileId}/export`, '_blank');
};

export default api;
