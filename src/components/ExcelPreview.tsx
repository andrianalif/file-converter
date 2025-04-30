import React, { useState, useEffect } from 'react';
import {
  Box,
  Paper,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Typography,
  Alert,
  Chip,
  CircularProgress,
  Select,
  MenuItem,
  FormControl,
  InputLabel,
  Grid,
  Card,
  CardContent,
  CardHeader,
  Divider,
} from '@mui/material';
import { read, utils } from 'xlsx';
import { ValidationError } from '../types';

interface ExcelPreviewProps {
  file: File | null;
  onValidationComplete: (errors: ValidationError[]) => void;
}

const ExcelPreview: React.FC<ExcelPreviewProps> = ({ file, onValidationComplete }) => {
  const [data, setData] = useState<any[][]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [loading, setLoading] = useState(false);
  const [errors, setErrors] = useState<ValidationError[]>([]);
  const [sheets, setSheets] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>('');

  useEffect(() => {
    if (file) {
      setLoading(true);
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const workbook = read(e.target?.result, { type: 'binary' });
          const sheetNames = workbook.SheetNames;
          setSheets(sheetNames);
          setSelectedSheet(sheetNames[0]);
          loadSheetData(workbook, sheetNames[0]);
        } catch (error) {
          console.error('Error reading Excel file:', error);
          setErrors([{
            type: 'format',
            message: 'Error reading Excel file',
            column: 'File'
          }]);
        } finally {
          setLoading(false);
        }
      };
      reader.readAsBinaryString(file);
    }
  }, [file]);

  const loadSheetData = (workbook: any, sheetName: string) => {
    try {
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
      
      if (jsonData.length > 0) {
        setHeaders(jsonData[0] as string[]);
        setData(jsonData.slice(1));
        validateData(jsonData);
      }
    } catch (error) {
      console.error('Error loading sheet data:', error);
      setErrors([{
        type: 'format',
        message: `Error loading sheet: ${sheetName}`,
        column: 'Sheet'
      }]);
    }
  };

  const handleSheetChange = (event: any) => {
    const newSheet = event.target.value;
    setSelectedSheet(newSheet);
    setLoading(true);
    const reader = new FileReader();
    reader.onload = (e) => {
      const workbook = read(e.target?.result, { type: 'binary' });
      loadSheetData(workbook, newSheet);
      setLoading(false);
    };
    reader.readAsBinaryString(file!);
  };

  const validateData = (jsonData: any[][]) => {
    const validationErrors: ValidationError[] = [];
    const headers = jsonData[0];
    const data = jsonData.slice(1);

    // Check for empty columns
    headers.forEach((header, colIndex) => {
      const isEmpty = data.every(row => !row[colIndex]);
      if (isEmpty) {
        validationErrors.push({
          type: 'empty',
          message: `Column "${header}" is empty`,
          column: header
        });
      }
    });

    // Check for format errors (example: numeric columns)
    headers.forEach((header, colIndex) => {
      const hasFormatError = data.some(row => {
        const value = row[colIndex];
        if (header.toLowerCase().includes('price') || 
            header.toLowerCase().includes('amount')) {
          return value && isNaN(Number(value));
        }
        return false;
      });

      if (hasFormatError) {
        validationErrors.push({
          type: 'format',
          message: `Invalid format in column "${header}"`,
          column: header
        });
      }
    });

    setErrors(validationErrors);
    onValidationComplete(validationErrors);
  };

  const getErrorTypeColor = (type: ValidationError['type'] | undefined) => {
    if (!type) return 'default';
    switch (type) {
      case 'empty':
        return 'warning';
      case 'format':
        return 'info';
      default:
        return 'default';
    }
  };

  if (loading) {
    return (
      <Box display="flex" justifyContent="center" alignItems="center" p={3}>
        <CircularProgress />
      </Box>
    );
  }

  if (!file || data.length === 0) {
    return null;
  }

  return (
    <Card sx={{ mt: 3, boxShadow: 3 }}>
      <CardHeader
        title="Excel Preview"
        subheader={`Sheet: ${selectedSheet}`}
        action={
          <FormControl sx={{ minWidth: 200, mr: 2 }}>
            <InputLabel id="sheet-select-label">Select Sheet</InputLabel>
            <Select
              labelId="sheet-select-label"
              value={selectedSheet}
              label="Select Sheet"
              onChange={handleSheetChange}
              sx={{ backgroundColor: 'background.paper' }}
            >
              {sheets.map((sheet) => (
                <MenuItem key={sheet} value={sheet}>
                  {sheet}
                </MenuItem>
              ))}
            </Select>
          </FormControl>
        }
      />
      <Divider />
      <CardContent>
        {errors.length > 0 && (
          <Box sx={{ mb: 2 }}>
            {errors.map((error, index) => (
              <Alert 
                key={index} 
                severity={getErrorTypeColor(error.type) as any}
                sx={{ mb: 1 }}
              >
                {error.message}
              </Alert>
            ))}
          </Box>
        )}

        <TableContainer component={Paper} sx={{ maxHeight: 400, boxShadow: 2 }}>
          <Table stickyHeader>
            <TableHead>
              <TableRow>
                {headers.map((header, index) => (
                  <TableCell key={index}>
                    <Box display="flex" alignItems="center" gap={1}>
                      {header}
                      {errors.some(e => e.column === header) && (
                        <Chip
                          size="small"
                          color={getErrorTypeColor(errors.find(e => e.column === header)?.type)}
                          label={errors.find(e => e.column === header)?.type || 'default'}
                        />
                      )}
                    </Box>
                  </TableCell>
                ))}
              </TableRow>
            </TableHead>
            <TableBody>
              {data.map((row, rowIndex) => (
                <TableRow key={rowIndex}>
                  {row.map((cell, cellIndex) => {
                    const columnError = errors.find(e => e.column === headers[cellIndex]);
                    const isErrorCell = columnError?.rows?.includes(rowIndex + 1);
                    
                    return (
                      <TableCell
                        key={cellIndex}
                        sx={{
                          backgroundColor: isErrorCell ? 'error.light' : 'inherit',
                          color: isErrorCell ? 'error.contrastText' : 'inherit'
                        }}
                      >
                        {cell}
                      </TableCell>
                    );
                  })}
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </TableContainer>
      </CardContent>
    </Card>
  );
};

export default ExcelPreview; 