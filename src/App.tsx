import React, { useState, useCallback } from 'react';
import { useDropzone } from 'react-dropzone';
import {
  Container,
  Box,
  Typography,
  Paper,
  Button,
  TextField,
  CircularProgress,
  Stepper,
  Step,
  StepLabel,
  Grid,
  Card,
  CardContent,
  IconButton,
  ThemeProvider,
  createTheme,
  CssBaseline,
  useMediaQuery,
  AppBar,
  Toolbar,
} from '@mui/material';
import { CloudUpload, Publish, Refresh, Description, Brightness4, Brightness7 } from '@mui/icons-material';
import axios from 'axios';
import StatusMessage from './components/StatusMessage';
import ExcelPreview from './components/ExcelPreview';
import { ValidationError } from './types';
import './App.css';

function App() {
  const prefersDarkMode = useMediaQuery('(prefers-color-scheme: dark)');
  const [isDarkMode, setIsDarkMode] = useState(prefersDarkMode);
  const [file, setFile] = useState<File | null>(null);
  const [title, setTitle] = useState('');
  const [isConverting, setIsConverting] = useState(false);
  const [isPublishing, setIsPublishing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);
  const [activeStep, setActiveStep] = useState(0);
  const [showStatus, setShowStatus] = useState(false);
  const [validationErrors, setValidationErrors] = useState<ValidationError[]>([]);

  const theme = createTheme({
    palette: {
      mode: isDarkMode ? 'dark' : 'light',
      primary: {
        main: '#4ec07d',
      },
      secondary: {
        main: '#ef8d9c',
      },
      background: {
        default: isDarkMode ? '#121212' : '#f5f5f5',
        paper: isDarkMode ? '#1e1e1e' : '#ffffff',
      },
    },
  });

  const toggleTheme = () => {
    setIsDarkMode(!isDarkMode);
  };

  const steps = ['Upload', 'Convert', 'Publish'];

  const onDrop = useCallback((acceptedFiles: File[]) => {
    if (acceptedFiles.length > 0) {
      const uploadedFile = acceptedFiles[0];
      setFile(uploadedFile);
      // Set title from filename (without extension)
      const fileName = uploadedFile.name.replace(/\.[^/.]+$/, "");
      setTitle(fileName);
      setActiveStep(0);
      setError(null);
      setSuccess(null);
      setShowStatus(false);
    }
  }, []);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']
    },
    multiple: false
  });

  const handleValidationComplete = (errors: ValidationError[]) => {
    setValidationErrors(errors);
  };

  const handleConvert = async () => {
    if (!file) {
      setError('Please select a file');
      setShowStatus(true);
      return;
    }

    if (validationErrors.length > 0) {
      setError('Please fix validation errors before converting');
      setShowStatus(true);
      return;
    }

    setIsConverting(true);
    setError(null);
    setSuccess(null);
    setShowStatus(false);

    try {
      const formData = new FormData();
      formData.append('file', file);
      formData.append('action', 'convert');

      const response = await axios.post('http://localhost:5000/api/process', formData, {
        headers: {
          'Content-Type': 'multipart/form-data'
        }
      });

      if (response.data.defaultTitle && !title) {
        setTitle(response.data.defaultTitle);
      }

      setIsConverting(false);
      setActiveStep(1);
      setSuccess('File converted successfully!');
      setShowStatus(true);
    } catch (err: any) {
      setError(err.response?.data?.error || 'Failed to convert file');
      setIsConverting(false);
      setShowStatus(true);
    }
  };

  const handlePublish = async () => {
    if (!file || !title) {
      setError('Please select a file and enter a title');
      setShowStatus(true);
      return;
    }

    setIsPublishing(true);
    setError(null);
    setSuccess(null);
    setShowStatus(false);

    try {
      const formData = new FormData();
      formData.append('file', file);
      formData.append('title', title);
      formData.append('action', 'publish');

      const response = await axios.post('http://localhost:5000/api/process', formData, {
        headers: {
          'Content-Type': 'multipart/form-data'
        }
      });

      if (response.data.defaultTitle && !title) {
        setTitle(response.data.defaultTitle);
      }

      if (response.data.url) {
        setSuccess(`File published successfully! URL: ${response.data.url}`);
      } else {
        setError('Publish successful but no URL returned');
      }
      setIsPublishing(false);
      setShowStatus(true);
    } catch (err: any) {
      setError(err.response?.data?.error || 'Failed to publish');
      setIsPublishing(false);
      setShowStatus(true);
    }
  };

  const handleCloseStatus = () => {
    setShowStatus(false);
    setError(null);
    setSuccess(null);
  };

  const handleReset = () => {
    setFile(null);
    setTitle('');
    setActiveStep(0);
    setError(null);
    setSuccess(null);
    setShowStatus(false);
  };

  return (
    <ThemeProvider theme={theme}>
      <CssBaseline />
      <AppBar position="static" color="default" elevation={1}>
        <Toolbar>
          <Box
            component="img"
            src="/images/logo.png"
            alt="VST ECS Logo"
            sx={{
              height: 40,
              mr: 2,
              objectFit: 'contain'
            }}
          />
          <Typography variant="h6" component="div" sx={{ flexGrow: 1 }}>
            Price List Publisher
          </Typography>
          <IconButton onClick={toggleTheme} color="inherit">
            {isDarkMode ? <Brightness7 /> : <Brightness4 />}
          </IconButton>
        </Toolbar>
      </AppBar>
      <Container maxWidth="md" sx={{ py: 4 }}>
        <Box sx={{ mb: 4, textAlign: 'center' }}>
          <Typography variant="h4" component="h1" gutterBottom sx={{ fontWeight: 'bold', color: 'primary.main' }}>
            Convert and Publish Your Price Lists
          </Typography>
          <Typography variant="subtitle1" color="text.secondary">
            Upload your Excel file and we'll help you convert it to a beautiful price list page
          </Typography>
        </Box>

        <Stepper activeStep={activeStep} sx={{ mb: 4 }}>
          {steps.map((label, index) => (
            <Step key={label} completed={activeStep > index}>
              <StepLabel>{label}</StepLabel>
            </Step>
          ))}
        </Stepper>

        <Grid container spacing={3}>
          <Grid item xs={12} md={6}>
            <Card elevation={3} sx={{ height: '100%' }}>
              <CardContent>
                <Typography variant="h6" gutterBottom>
                  Upload Excel File
                </Typography>
                <Paper
                  {...getRootProps()}
                  sx={{
                    p: 3,
                    textAlign: 'center',
                    cursor: 'pointer',
                    backgroundColor: isDragActive ? 'primary.light' : 'background.paper',
                    border: '2px dashed',
                    borderColor: isDragActive ? 'primary.main' : 'divider',
                    transition: 'all 0.3s ease',
                    '&:hover': {
                      backgroundColor: 'primary.light',
                      borderColor: 'primary.main',
                    }
                  }}
                >
                  <input {...getInputProps()} />
                  <CloudUpload sx={{ fontSize: 48, color: 'primary.main', mb: 2 }} />
                  {isDragActive ? (
                    <Typography color="primary">Drop the file here...</Typography>
                  ) : (
                    <Typography>
                      Drag and drop a file here, or click to select a file
                    </Typography>
                  )}
                </Paper>
                {file && (
                  <Box sx={{ mt: 2, display: 'flex', alignItems: 'center', gap: 1 }}>
                    <Description color="primary" />
                    <Typography variant="body2">{file.name}</Typography>
                  </Box>
                )}
              </CardContent>
            </Card>
          </Grid>

          <Grid item xs={12} md={6}>
            <Card elevation={3} sx={{ height: '100%' }}>
              <CardContent>
                <Typography variant="h6" gutterBottom>
                  Page Details
                </Typography>
                <TextField
                  fullWidth
                  label="Page Title"
                  value={title}
                  onChange={(e) => setTitle(e.target.value)}
                  sx={{ mb: 2 }}
                  placeholder="Enter the title for your price list page"
                />
                <Box sx={{ display: 'flex', gap: 2, flexWrap: 'wrap' }}>
                  <Button
                    variant="contained"
                    color="primary"
                    onClick={handleConvert}
                    disabled={!file || isConverting || activeStep > 0 || validationErrors.length > 0}
                    startIcon={isConverting ? <CircularProgress size={20} /> : <Refresh />}
                  >
                    {isConverting ? 'Converting...' : 'Convert'}
                  </Button>

                  <Button
                    variant="contained"
                    color="success"
                    onClick={handlePublish}
                    disabled={!file || !title || isPublishing || activeStep < 1}
                    startIcon={isPublishing ? <CircularProgress size={20} /> : <Publish />}
                  >
                    {isPublishing ? 'Publishing...' : 'Publish'}
                  </Button>

                  <Button
                    variant="outlined"
                    color="secondary"
                    onClick={handleReset}
                    disabled={!file && !title}
                  >
                    Reset
                  </Button>
                </Box>
              </CardContent>
            </Card>
          </Grid>
        </Grid>

        {file && (
          <ExcelPreview
            file={file}
            onValidationComplete={handleValidationComplete}
          />
        )}

        {showStatus && (
          <StatusMessage
            type={error ? 'error' : 'success'}
            title={error ? 'Error!' : 'Success!'}
            message={error || success || ''}
            buttonText={error ? 'Try Again' : 'Continue'}
            onButtonClick={handleCloseStatus}
          />
        )}
      </Container>
    </ThemeProvider>
  );
}

export default App; 