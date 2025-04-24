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
  Alert,
  Snackbar,
  Stepper,
  Step,
  StepLabel,
  Grid,
  Card,
  CardContent,
  Divider
} from '@mui/material';
import { CloudUpload, Publish, Refresh, Description } from '@mui/icons-material';
import axios from 'axios';
import './App.css';

function App() {
  const [file, setFile] = useState<File | null>(null);
  const [title, setTitle] = useState('');
  const [isConverting, setIsConverting] = useState(false);
  const [isPublishing, setIsPublishing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);
  const [activeStep, setActiveStep] = useState(0);

  const steps = ['Upload', 'Convert', 'Publish'];

  const onDrop = useCallback((acceptedFiles: File[]) => {
    if (acceptedFiles.length > 0) {
      setFile(acceptedFiles[0]);
      setActiveStep(0);
      setError(null);
      setSuccess(null);
    }
  }, []);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']
    },
    multiple: false
  });

  const handleConvert = async () => {
    if (!file) {
      setError('Please select a file');
      return;
    }

    setIsConverting(true);
    setError(null);
    setSuccess(null);

    try {
      const formData = new FormData();
      formData.append('file', file);
      formData.append('action', 'convert');

      await axios.post('http://localhost:5000/api/process', formData, {
        headers: {
          'Content-Type': 'multipart/form-data'
        }
      });

      setIsConverting(false);
      setActiveStep(1);
      setSuccess('File converted successfully!');
    } catch (err: any) {
      setError(err.response?.data?.error || 'Failed to convert file');
      setIsConverting(false);
    }
  };

  const handlePublish = async () => {
    if (!file || !title) {
      setError('Please select a file and enter a title');
      return;
    }

    setIsPublishing(true);
    setError(null);
    setSuccess(null);

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

      if (response.data.url) {
        setSuccess(`File published successfully! URL: ${response.data.url}`);
      } else {
        setError('Publish successful but no URL returned');
      }
      setIsPublishing(false);
    } catch (err: any) {
      setError(err.response?.data?.error || 'Failed to publish');
      setIsPublishing(false);
    }
  };

  const handleCloseSnackbar = () => {
    setError(null);
    setSuccess(null);
  };

  const handleReset = () => {
    setFile(null);
    setTitle('');
    setActiveStep(0);
    setError(null);
    setSuccess(null);
  };

  return (
    <Container maxWidth="md" sx={{ py: 4 }}>
      <Box sx={{ mb: 4, textAlign: 'center' }}>
        <Typography variant="h3" component="h1" gutterBottom sx={{ fontWeight: 'bold', color: 'primary.main' }}>
          Price List Publisher
        </Typography>
        <Typography variant="subtitle1" color="text.secondary">
          Convert and publish your Excel price lists to WordPress
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
                  disabled={!file || isConverting || activeStep > 0}
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

      <Snackbar
        open={!!error || !!success}
        autoHideDuration={6000}
        onClose={handleCloseSnackbar}
        anchorOrigin={{ vertical: 'bottom', horizontal: 'center' }}
      >
        <Alert
          onClose={handleCloseSnackbar}
          severity={error ? 'error' : 'success'}
          sx={{ width: '100%' }}
        >
          {error || success}
        </Alert>
      </Snackbar>
    </Container>
  );
}

export default App; 