import React, { useState } from 'react';
import { 
  Container, 
  Box, 
  Typography, 
  Paper, 
  Button, 
  TextField,
  CircularProgress,
  Alert,
  Snackbar
} from '@mui/material';
import { useDropzone } from 'react-dropzone';
import axios from 'axios';
import './App.css';

function App() {
  const [file, setFile] = useState<File | null>(null);
  const [title, setTitle] = useState('');
  const [htmlContent, setHtmlContent] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState('');

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'],
      'application/pdf': ['.pdf']
    },
    maxFiles: 1,
    onDrop: (acceptedFiles) => {
      setFile(acceptedFiles[0]);
      setTitle(acceptedFiles[0].name.split('.')[0]);
    }
  });

  const handleConvert = async () => {
    if (!file) {
      setError('Please select a file first');
      return;
    }

    setLoading(true);
    setError('');
    setSuccess('');

    const formData = new FormData();
    formData.append('file', file);

    try {
      const response = await axios.post('http://localhost:5000/api/convert', formData);
      setHtmlContent(response.data.html);
      setSuccess('File converted successfully!');
    } catch (err) {
      setError('Failed to convert file. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  const handlePublish = async () => {
    if (!htmlContent || !title) {
      setError('Please convert a file first and provide a title');
      return;
    }

    setLoading(true);
    setError('');
    setSuccess('');

    try {
      const response = await axios.post('http://localhost:5000/api/publish', {
        title,
        content: htmlContent
      });
      setSuccess(`Page published successfully! URL: ${response.data.url}`);
    } catch (err) {
      setError('Failed to publish page. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <Container maxWidth="lg">
      <Box sx={{ my: 4 }}>
        <Typography variant="h3" component="h1" gutterBottom align="center">
          File Converter & Publisher
        </Typography>
        
        <Paper 
          {...getRootProps()} 
          sx={{ 
            p: 4, 
            mb: 4, 
            textAlign: 'center',
            border: '2px dashed #ccc',
            backgroundColor: isDragActive ? '#f5f5f5' : 'white',
            cursor: 'pointer'
          }}
        >
          <input {...getInputProps()} />
          {isDragActive ? (
            <Typography>Drop the file here ...</Typography>
          ) : (
            <Typography>
              Drag and drop a file here, or click to select a file
              <br />
              (Supports Excel, Word, and PDF files)
            </Typography>
          )}
        </Paper>

        {file && (
          <Paper sx={{ p: 4, mb: 4 }}>
            <Typography variant="h6" gutterBottom>
              Selected File: {file.name}
            </Typography>
            <TextField
              fullWidth
              label="Page Title"
              value={title}
              onChange={(e) => setTitle(e.target.value)}
              sx={{ mb: 2 }}
            />
            <Box sx={{ display: 'flex', gap: 2 }}>
              <Button
                variant="contained"
                onClick={handleConvert}
                disabled={loading}
                sx={{ flex: 1 }}
              >
                {loading ? <CircularProgress size={24} /> : 'Convert'}
              </Button>
              <Button
                variant="contained"
                color="secondary"
                onClick={handlePublish}
                disabled={loading || !htmlContent}
                sx={{ flex: 1 }}
              >
                Publish to WordPress
              </Button>
            </Box>
          </Paper>
        )}

        {htmlContent && (
          <Paper sx={{ p: 4 }}>
            <Typography variant="h6" gutterBottom>
              Preview
            </Typography>
            <div dangerouslySetInnerHTML={{ __html: htmlContent }} />
          </Paper>
        )}

        <Snackbar 
          open={!!error} 
          autoHideDuration={6000} 
          onClose={() => setError('')}
        >
          <Alert severity="error" onClose={() => setError('')}>
            {error}
          </Alert>
        </Snackbar>

        <Snackbar 
          open={!!success} 
          autoHideDuration={6000} 
          onClose={() => setSuccess('')}
        >
          <Alert severity="success" onClose={() => setSuccess('')}>
            {success}
          </Alert>
        </Snackbar>
      </Box>
    </Container>
  );
}

export default App; 