import React from 'react';
import {
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
  List,
  ListItem,
  ListItemText,
  ListItemSecondaryAction,
  IconButton,
  Typography,
  Box,
  Divider,
} from '@mui/material';
import { History as HistoryIcon, Delete as DeleteIcon, OpenInNew as OpenInNewIcon } from '@mui/icons-material';

interface HistoryItem {
  id: string;
  title: string;
  date: string;
  url?: string;
}

interface HistoryDialogProps {
  open: boolean;
  onClose: () => void;
  history: HistoryItem[];
  onDelete: (id: string) => void;
}

const HistoryDialog: React.FC<HistoryDialogProps> = ({ open, onClose, history, onDelete }) => {
  return (
    <Dialog open={open} onClose={onClose} maxWidth="md" fullWidth>
      <DialogTitle>
        <Box display="flex" alignItems="center">
          <HistoryIcon sx={{ mr: 1 }} />
          <Typography variant="h6">Conversion History</Typography>
        </Box>
      </DialogTitle>
      <DialogContent>
        {history.length === 0 ? (
          <Typography variant="body1" color="text.secondary" sx={{ textAlign: 'center', py: 2 }}>
            No conversion history available
          </Typography>
        ) : (
          <List>
            {history.map((item, index) => (
              <React.Fragment key={item.id}>
                <ListItem>
                  <ListItemText
                    primary={item.title}
                    secondary={new Date(item.date).toLocaleString()}
                  />
                  <ListItemSecondaryAction>
                    {item.url && (
                      <IconButton
                        edge="end"
                        aria-label="open"
                        onClick={() => window.open(item.url, '_blank')}
                        sx={{ mr: 1 }}
                      >
                        <OpenInNewIcon />
                      </IconButton>
                    )}
                    <IconButton
                      edge="end"
                      aria-label="delete"
                      onClick={() => onDelete(item.id)}
                    >
                      <DeleteIcon />
                    </IconButton>
                  </ListItemSecondaryAction>
                </ListItem>
                {index < history.length - 1 && <Divider />}
              </React.Fragment>
            ))}
          </List>
        )}
      </DialogContent>
      <DialogActions>
        <Button onClick={onClose}>Close</Button>
      </DialogActions>
    </Dialog>
  );
};

export default HistoryDialog; 