
import React, { useState } from 'react';
import { ApiExcelData } from '@/services/api';
import FileUpload from '@/components/FileUpload';
import SpreadsheetEditor from '@/components/SpreadsheetEditor';

const Index = () => {
  const [file, setFile] = useState<File | null>(null);
  const [excelData, setExcelData] = useState<ApiExcelData | null>(null);
  const [isLoading, setIsLoading] = useState(false);

  const handleFileLoaded = (file: File, apiData: ApiExcelData) => {
    setFile(file);
    setExcelData(apiData);
    setIsLoading(false);
  };

  const handleDataChange = (newData: ApiExcelData) => {
    setExcelData(newData);
  };

  return (
    <div className="container mx-auto p-4 md:p-8 max-w-7xl h-full flex flex-col">
      <h1 className="text-3xl font-bold text-center mb-8">Excel Editor</h1>
      
      {!excelData ? (
        <div className="flex-grow flex flex-col items-center justify-center">
          <FileUpload onFileLoaded={handleFileLoaded} isLoading={isLoading} />
          <p className="mt-8 text-muted-foreground text-center max-w-lg">
            Upload an Excel file to view and edit its contents. Your data will be processed securely.
          </p>
        </div>
      ) : (
        <div className="flex-grow flex flex-col h-[calc(100vh-12rem)]">
          <SpreadsheetEditor 
            excelData={excelData}
            onDataChange={handleDataChange}
          />
        </div>
      )}
    </div>
  );
};

export default Index;
