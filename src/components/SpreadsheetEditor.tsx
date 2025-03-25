
import React, { useState, useEffect, useCallback } from 'react';
import { MaterialReactTable, type MRT_ColumnDef } from 'material-react-table';
import { ExcelData, updateCellValue, exportToExcel } from '@/utils/excelUtils';
import { Button } from '@/components/ui/button';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { toast } from '@/hooks/use-toast';
import { Download, Save } from 'lucide-react';
import { ThemeProvider, createTheme } from '@mui/material';

interface SpreadsheetEditorProps {
  excelData: ExcelData;
  fileName: string;
  onDataChange: (newData: ExcelData) => void;
}

const SpreadsheetEditor: React.FC<SpreadsheetEditorProps> = ({
  excelData,
  fileName,
  onDataChange,
}) => {
  const [activeSheet, setActiveSheet] = useState<string>(excelData.activeSheet);
  const [columns, setColumns] = useState<MRT_ColumnDef<any>[]>([]);
  
  // Create the Material-UI theme to match our design
  const theme = createTheme({
    palette: {
      primary: {
        main: 'hsl(200, 70%, 50%)',
      },
      background: {
        default: 'transparent',
      },
    },
    typography: {
      fontFamily: [
        '-apple-system',
        'BlinkMacSystemFont',
        'San Francisco',
        'Segoe UI',
        'Roboto',
        'Helvetica Neue',
        'sans-serif',
      ].join(','),
    },
    components: {
      MuiPaper: {
        styleOverrides: {
          root: {
            backgroundColor: 'transparent',
            backgroundImage: 'none',
            boxShadow: 'none',
          },
        },
      },
      MuiTableCell: {
        styleOverrides: {
          root: {
            borderBottom: '1px solid hsla(var(--border) / 0.5)',
            fontSize: '0.875rem',
          },
          head: {
            backgroundColor: 'hsla(var(--secondary) / 0.8)',
            backdropFilter: 'blur(2px)',
            position: 'sticky',
            top: 0,
            zIndex: 10,
            fontWeight: 500,
            fontSize: '0.875rem',
            padding: '12px 16px',
          },
        },
      },
      MuiTableRow: {
        styleOverrides: {
          root: {
            '&:hover': {
              backgroundColor: 'hsla(var(--muted) / 0.3)',
            },
          },
        },
      },
      MuiInputBase: {
        styleOverrides: {
          root: {
            fontSize: '0.875rem',
          },
        },
      },
    },
  });

  // Setup columns based on headers
  useEffect(() => {
    if (excelData && excelData.headers) {
      const tableColumns = excelData.headers.map((header, index) => ({
        accessorKey: index.toString(),
        header: header || `Column ${index + 1}`,
        size: 180,
        muiTableHeadCellProps: {
          align: 'left',
        },
        muiTableBodyCellProps: {
          align: 'left',
        },
      }));
      setColumns(tableColumns);
    }
  }, [excelData]);

  // Transform the data for the table
  const tableData = excelData.data.map((row, rowIdx) => {
    const rowData: Record<string, any> = {};
    row.forEach((cell, cellIdx) => {
      rowData[cellIdx.toString()] = cell;
    });
    return rowData;
  });

  // Handle cell edit
  const handleCellEdit = useCallback(
    (rowIndex: number, columnId: string, value: any) => {
      const columnIndex = parseInt(columnId);
      const newData = updateCellValue(excelData, rowIndex, columnIndex, value);
      onDataChange(newData);
    },
    [excelData, onDataChange]
  );

  // Handle export
  const handleExport = () => {
    exportToExcel(excelData, fileName);
  };

  // Handle sheet change
  const handleSheetChange = (sheetName: string) => {
    setActiveSheet(sheetName);
    toast({
      title: 'Sheet changed',
      description: `Switched to ${sheetName}. Note: Multi-sheet editing is not supported in this version.`,
    });
  };

  return (
    <div className="w-full h-full flex flex-col gap-4 bg-card/30 backdrop-blur-xs rounded-xl border border-border/50 shadow-card animate-scale-in overflow-hidden">
      <div className="flex items-center justify-between p-4 border-b border-border/50">
        <h2 className="text-lg font-medium">{fileName}</h2>
        <div className="flex items-center gap-2">
          <Button
            variant="outline"
            size="sm"
            className="gap-1.5"
            onClick={handleExport}
          >
            <Download className="w-4 h-4" />
            <span>Export</span>
          </Button>
        </div>
      </div>

      {excelData.sheetNames.length > 1 && (
        <div className="px-4">
          <Tabs defaultValue={activeSheet} onValueChange={handleSheetChange}>
            <TabsList className="w-full h-10 overflow-x-auto max-w-full flex-nowrap">
              {excelData.sheetNames.map((sheet) => (
                <TabsTrigger
                  key={sheet}
                  value={sheet}
                  className="whitespace-nowrap"
                >
                  {sheet}
                </TabsTrigger>
              ))}
            </TabsList>
          </Tabs>
        </div>
      )}

      <div className="flex-1 overflow-hidden">
        <ThemeProvider theme={theme}>
          <MaterialReactTable
            columns={columns}
            data={tableData}
            enableColumnActions={false}
            enableColumnFilters={false}
            enablePagination={false}
            enableBottomToolbar={false}
            enableTopToolbar={false}
            enableColumnResizing
            enableEditing
            editingMode="cell"
            muiTableProps={{
              sx: {
                tableLayout: 'fixed',
              },
            }}
            muiTableContainerProps={{
              sx: {
                maxHeight: '100%',
                height: '100%',
              },
            }}
            onEditingRowSave={({ row, values }) => {
              Object.entries(values).forEach(([columnId, value]) => {
                handleCellEdit(row.index, columnId, value);
              });
            }}
            muiEditTextFieldProps={({ cell }) => ({
              onBlur: (event) => {
                handleCellEdit(
                  cell.row.index,
                  cell.column.id,
                  event.target.value
                );
              },
            })}
          />
        </ThemeProvider>
      </div>
    </div>
  );
};

export default SpreadsheetEditor;
