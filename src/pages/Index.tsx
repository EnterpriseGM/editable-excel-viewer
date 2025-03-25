
import React, { useState } from 'react';
import FileUpload from '@/components/FileUpload';
import SpreadsheetEditor from '@/components/SpreadsheetEditor';
import { parseExcelFile, ExcelData } from '@/utils/excelUtils';
import { toast } from '@/hooks/use-toast';

const Index = () => {
  const [excelData, setExcelData] = useState<ExcelData | null>(null);
  const [fileName, setFileName] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(false);

  const handleFileLoaded = async (file: File) => {
    setIsLoading(true);
    try {
      const data = await parseExcelFile(file);
      if (data) {
        setExcelData(data);
        setFileName(file.name);
        toast({
          title: 'File loaded successfully',
          description: `Loaded ${file.name} with ${data.data.length} rows and ${data.headers.length} columns.`,
        });
      }
    } catch (error) {
      console.error('Error loading file:', error);
      toast({
        title: 'Error',
        description: 'Failed to load the file. Please try again.',
        variant: 'destructive',
      });
    } finally {
      setIsLoading(false);
    }
  };

  const handleDataChange = (newData: ExcelData) => {
    setExcelData(newData);
  };

  return (
    <div className="min-h-screen w-full flex flex-col items-center bg-background p-4 md:p-8 animate-fade-in">
      <main className="container mx-auto w-full max-w-7xl flex flex-col gap-8">
        <header className="w-full text-center mb-6 animate-slide-up">
          <h1 className="text-3xl md:text-4xl font-semibold tracking-tight mb-2">
            Excel Viewer & Editor
          </h1>
          <p className="text-muted-foreground max-w-2xl mx-auto">
            Upload your Excel file to view and edit it directly in the browser. Changes can be exported back to Excel format.
          </p>
        </header>

        <div className="w-full flex-1 flex flex-col gap-8">
          {!excelData ? (
            <FileUpload onFileLoaded={handleFileLoaded} isLoading={isLoading} />
          ) : (
            <div className="flex flex-col gap-6 h-[calc(100vh-200px)]">
              <div className="flex items-center justify-between">
                <button
                  onClick={() => setExcelData(null)}
                  className="text-sm text-primary hover:underline transition-all"
                >
                  ← Upload another file
                </button>
              </div>
              <SpreadsheetEditor
                excelData={excelData}
                fileName={fileName}
                onDataChange={handleDataChange}
              />
            </div>
          )}
        </div>
      </main>

      <footer className="w-full text-center mt-auto pt-8 text-sm text-muted-foreground">
        <p>Excel Viewer & Editor • Built with modern web technologies</p>
      </footer>
    </div>
  );
};

export default Index;
