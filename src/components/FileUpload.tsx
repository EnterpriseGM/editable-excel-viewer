
import React, { useCallback } from 'react';
import { useDropzone } from 'react-dropzone';
import { cn } from '@/lib/utils';
import { ArrowUp, FileSpreadsheet } from 'lucide-react';
import { toast } from '@/hooks/use-toast';
import { uploadExcelFile } from '@/services/api';

interface FileUploadProps {
  onFileLoaded: (file: File, apiData: any) => void;
  isLoading?: boolean;
}

const FileUpload: React.FC<FileUploadProps> = ({ onFileLoaded, isLoading = false }) => {
  const onDrop = useCallback(
    async (acceptedFiles: File[]) => {
      if (acceptedFiles.length === 0) return;
      
      const file = acceptedFiles[0];
      
      // Check if the file is an Excel file
      const isExcelFile = 
        file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
        file.type === 'application/vnd.ms-excel' ||
        file.name.endsWith('.xlsx') ||
        file.name.endsWith('.xls');
      
      if (!isExcelFile) {
        toast({
          title: 'Invalid file',
          description: 'Please upload an Excel file (.xlsx or .xls)',
          variant: 'destructive',
        });
        return;
      }
      
      try {
        const apiData = await uploadExcelFile(file);
        onFileLoaded(file, apiData);
      } catch (error) {
        toast({
          title: 'Error',
          description: 'Failed to upload the file. Please try again.',
          variant: 'destructive',
        });
      }
    },
    [onFileLoaded]
  );

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls'],
    },
    multiple: false,
    disabled: isLoading,
  });

  return (
    <div
      {...getRootProps()}
      className={cn(
        'w-full h-60 flex flex-col items-center justify-center rounded-xl p-6 custom-transition cursor-pointer',
        'border-2 border-dashed border-primary/30 hover:border-primary/60',
        'bg-accent/50 hover:bg-accent/80 backdrop-blur-xs',
        'animate-fade-in duration-300 ease-out',
        isDragActive && 'border-primary bg-accent',
        isLoading && 'opacity-70 cursor-wait'
      )}
    >
      <input {...getInputProps()} data-testid="file-input" />
      
      <div className="flex flex-col items-center gap-4 text-center">
        {isDragActive ? (
          <>
            <div className="w-16 h-16 rounded-full bg-primary/10 flex items-center justify-center mb-2 animate-pulse">
              <ArrowUp className="w-8 h-8 text-primary" />
            </div>
            <p className="text-lg font-medium text-primary">Drop your Excel file here</p>
          </>
        ) : (
          <>
            <div className="w-16 h-16 rounded-full bg-primary/10 flex items-center justify-center mb-2 group-hover:bg-primary/20 custom-transition">
              <FileSpreadsheet className="w-8 h-8 text-primary" />
            </div>
            <p className="text-lg font-medium">Drag & drop your Excel file</p>
            <p className="text-sm text-muted-foreground">
              or click to browse (supports .xlsx and .xls)
            </p>
          </>
        )}
      </div>
    </div>
  );
};

export default FileUpload;
