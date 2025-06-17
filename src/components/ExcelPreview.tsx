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
  Accordion,
  AccordionSummary,
  AccordionDetails,
} from '@mui/material';
import ExpandMoreIcon from '@mui/icons-material/ExpandMore';
import { read, utils } from 'xlsx';
import { ValidationError } from '../types';

interface ExcelPreviewProps {
  file: File | null;
  onValidationComplete: (errors: ValidationError[]) => void;
  contextData?: Record<string, any[]>;
  isConverted?: boolean;
  selectedSheet?: string;
}

interface ContextData {
  product_name: string;
  product_number: string;
  description: string;
  category: string;
  warranty: string;
  price: number | null;
  metadata: {
    is_subscription: boolean;
    is_service: boolean;
    specifications: Record<string, any>;
  };
}

const ExcelPreview: React.FC<ExcelPreviewProps> = ({ file, onValidationComplete, contextData, isConverted, selectedSheet }) => {
  const [data, setData] = useState<any[][]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [loading, setLoading] = useState(false);
  const [errors, setErrors] = useState<ValidationError[]>([]);
  const [sheets, setSheets] = useState<string[]>([]);

  useEffect(() => {
    if (file) {
      setLoading(true);
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const workbook = read(e.target?.result, { type: 'binary' });
          const sheetNames = workbook.SheetNames;
          setSheets(sheetNames);
          setHeaders([]);
          setData([]);
          const sheetToLoad = selectedSheet || sheetNames[0];
          await loadSheetData(workbook, sheetToLoad);
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
  }, [file, selectedSheet]);

  const generateContexts = (data: any[][], headers: string[], sheetName: string): ContextData[] => {
    const contexts: ContextData[] = [];
    
    for (const row of data) {
      if (row.every(cell => !cell)) continue; // Skip empty rows
      
      const rowData: Record<string, any> = {};
      headers.forEach((header, index) => {
        if (header) {
          rowData[header] = row[index];
        }
      });
      
      // Extract product information
      const productNumber = rowData['Product Number'] || '';
      const description = rowData['Description'] || '';
      const price = rowData['Price'] || null;
      const warranty = rowData['Warranty'] || '';
      
      // Generate context
      const context: ContextData = {
        product_name: description.split(' - ')[0] || productNumber,
        product_number: productNumber,
        description: description,
        category: sheetName,
        warranty: warranty,
        price: price,
        metadata: {
          is_subscription: description.toLowerCase().includes('subscription'),
          is_service: description.toLowerCase().includes('service'),
          specifications: extractSpecifications(description)
        }
      };
      
      contexts.push(context);
    }
    
    return contexts;
  };

  const extractSpecifications = (description: string): Record<string, any> => {
    const specs: Record<string, any> = {};
    
    // Common specification patterns
    const patterns = {
      processor: /(?:Intel|AMD|Core|Ryzen|i\d|i\d-\d{4}[A-Z]?)/i,
      ram: /(\d+GB(?:\s+RAM)?)/i,
      storage: /(\d+GB(?:\s+SSD|\s+HDD)?)/i,
      display: /(\d+(?:\.\d+)?["\'](?:\s+FHD|\s+UHD|\s+4K)?)/i,
      os: /(Windows\s+\d+(?:\s+Pro)?|Linux|macOS)/i,
      warranty: /(\d+\s+Year(?:\s+on-site)?)/i
    };
    
    for (const [key, pattern] of Object.entries(patterns)) {
      const match = description.match(pattern);
      if (match) {
        specs[key] = match[0];
      }
    }
    
    // Extract features
    const features = [];
    const featureKeywords = ['Fingerprint', 'Backlit', 'Bluetooth', 'Wi-Fi', 'USB', 'HDMI', 'DisplayPort', 'Thunderbolt'];
    for (const keyword of featureKeywords) {
      if (description.toLowerCase().includes(keyword.toLowerCase())) {
        features.push(keyword);
      }
    }
    
    if (features.length > 0) {
      specs['features'] = features;
    }
    
    return specs;
  };

  const loadSheetData = async (workbook: any, sheetName: string) => {
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

  function getImportantContextPoints(contextArr: any[]): string[] {
    if (!contextArr || contextArr.length === 0) return [];
    const first = contextArr[0];
    const points: string[] = [];
    if (first.product_name) points.push(`Product Name: "${first.product_name}"`);
    if (first.product_number) points.push(`Product Number: "${first.product_number}"`);
    if (first.category) points.push(`Category: "${first.category}"`);
    if (first.price !== undefined && first.price !== null && first.price !== '') points.push(`Price: "${first.price}"`);
    if (first.warranty) points.push(`Warranty: "${first.warranty}"`);
    if (first.description) points.push(`Description: "${first.description}"`);
    if (first.metadata && first.metadata.column_analysis && first.metadata.column_analysis.important_columns) {
      points.push(`Important Columns: ${first.metadata.column_analysis.important_columns.join(', ')}`);
    }
    return points;
  }

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
    <Box sx={{ mt: 3 }}>
      <Card sx={{ mb: 3, boxShadow: 3 }}>
        <CardHeader
          title="Excel Preview"
          subheader={`Sheet: ${selectedSheet}`}
          action={
            <FormControl sx={{ minWidth: 200, mr: 2 }}>
              <InputLabel id="sheet-select-label">Select Sheet</InputLabel>
              <Select
                labelId="sheet-select-label"
                value={selectedSheet || ''}
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

      {isConverted && contextData && selectedSheet && contextData[selectedSheet] && (
        <Card sx={{ boxShadow: 3 }}>
          <CardHeader
            title="Context Summary"
            subheader={`Summary of extracted contexts for ${selectedSheet}`}
          />
          <Divider />
          <CardContent>
            <Box sx={{ mb: 3, p: 2, bgcolor: 'info.light', borderRadius: 1 }}>
              <Typography variant="h6" gutterBottom>Sheet Summary</Typography>
              <Grid container spacing={2}>
                <Grid item xs={12} md={4}>
                  <Typography variant="subtitle2">Total Products</Typography>
                  <Typography variant="body1">{contextData[selectedSheet].length}</Typography>
                </Grid>
                <Grid item xs={12} md={4}>
                  <Typography variant="subtitle2">Categories</Typography>
                  <Typography variant="body1">
                    {Array.from(new Set(contextData[selectedSheet].map((ctx: any) => ctx.category))).join(', ')}
                  </Typography>
                </Grid>
                <Grid item xs={12} md={4}>
                  <Typography variant="subtitle2">Columns Used</Typography>
                  <Typography variant="body1">{headers.join(', ')}</Typography>
                </Grid>
              </Grid>
              <Box sx={{ mt: 2 }}>
                <Typography variant="subtitle2" gutterBottom>Important Context Points:</Typography>
                <ul>
                  {getImportantContextPoints(contextData[selectedSheet]).map((point, idx) => (
                    <li key={idx}>{point}</li>
                  ))}
                </ul>
              </Box>
            </Box>

            {contextData[selectedSheet].map((context: any, index: number) => (
              <Accordion key={index} sx={{ mb: 2 }}>
                <AccordionSummary expandIcon={<ExpandMoreIcon />}>
                  <Typography variant="subtitle1">{context.product_name}</Typography>
                </AccordionSummary>
                <AccordionDetails>
                  <Grid container spacing={2}>
                    <Grid item xs={12} md={6}>
                      <Typography variant="subtitle2">Product Details</Typography>
                      <Box component="pre" sx={{ 
                        p: 2, 
                        bgcolor: 'grey.100', 
                        borderRadius: 1,
                        overflow: 'auto',
                        fontSize: '0.875rem'
                      }}>
                        {JSON.stringify({
                          product_number: context.product_number,
                          description: context.description,
                          category: context.category,
                          warranty: context.warranty,
                          price: context.price
                        }, null, 2)}
                      </Box>
                    </Grid>
                    <Grid item xs={12} md={6}>
                      <Typography variant="subtitle2">Specifications</Typography>
                      <Box component="pre" sx={{ 
                        p: 2, 
                        bgcolor: 'grey.100', 
                        borderRadius: 1,
                        overflow: 'auto',
                        fontSize: '0.875rem'
                      }}>
                        {JSON.stringify(context.metadata.specifications, null, 2)}
                      </Box>
                    </Grid>
                  </Grid>
                </AccordionDetails>
              </Accordion>
            ))}
          </CardContent>
        </Card>
      )}
    </Box>
  );
};

export default ExcelPreview; 