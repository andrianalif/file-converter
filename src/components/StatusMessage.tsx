import React from 'react';
import { Box, Typography, Button } from '@mui/material';
import { styled } from '@mui/material/styles';

interface MessageBoxProps {
  messagetype: 'success' | 'error';
}

interface MouthProps {
  messagetype: 'success' | 'error';
}

interface ShadowProps {
  messagetype: 'success' | 'error';
}

interface ActionButtonProps {
  messagetype: 'success' | 'error';
}

const Overlay = styled(Box)(({ theme }) => ({
  position: 'fixed',
  top: 0,
  left: 0,
  right: 0,
  bottom: 0,
  backgroundColor: 'rgba(0, 0, 0, 0.5)',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  zIndex: 1000,
  padding: 0,
  margin: 0,
}));

const Container = styled(Box)(({ theme }) => ({
  position: 'fixed',
  top: '50%',
  left: '50%',
  transform: 'translate(-50%, -50%)',
  width: '300px',
  aspectRatio: '1',
  zIndex: 1001,
}));

const MessageBox = styled(Box)<MessageBoxProps>(({ theme, messagetype }) => ({
  width: '100%',
  height: '100%',
  background: messagetype === 'success' 
    ? 'linear-gradient(to bottom right, #b0db7d 40%, #99dbb4 100%)'
    : 'linear-gradient(to bottom left, #ef8d9c 40%, #ffc39e 100%)',
  borderRadius: '20px',
  boxShadow: theme.palette.mode === 'dark' 
    ? '0 8px 32px rgba(0, 0, 0, 0.3)' 
    : '5px 5px 20px rgba(203, 205, 211, 0.1)',
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'center',
  justifyContent: 'center',
  padding: '24px',
  gap: '16px',
  perspective: '40px',
}));

const Face = styled(Box)(({ theme }) => ({
  width: '22%',
  height: '22%',
  background: theme.palette.mode === 'dark' ? '#2f2f2f' : '#fcfcfc',
  borderRadius: '50%',
  border: `1px solid ${theme.palette.mode === 'dark' ? '#555' : '#777'}`,
  position: 'relative',
  marginBottom: '20px',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  animation: 'bounce 1s ease-in infinite',
  '@keyframes bounce': {
    '50%': {
      transform: 'translateY(-10px)',
    }
  }
}));

const Eye = styled(Box)(({ theme }) => ({
  width: '5px',
  height: '5px',
  background: theme.palette.mode === 'dark' ? '#999' : '#777',
  borderRadius: '50%',
  position: 'absolute',
  top: '40%',
  '&.left': {
    left: '20%',
  },
  '&.right': {
    left: '68%',
  },
}));

const Mouth = styled(Box)<MouthProps>(({ theme, messagetype }) => ({
  position: 'absolute',
  width: '7px',
  height: '7px',
  borderRadius: '50%',
  border: '2px solid',
  borderColor: messagetype === 'success' 
    ? `transparent ${theme.palette.mode === 'dark' ? '#999' : '#777'} ${theme.palette.mode === 'dark' ? '#999' : '#777'} transparent`
    : `${theme.palette.mode === 'dark' ? '#999' : '#777'} transparent transparent ${theme.palette.mode === 'dark' ? '#999' : '#777'}`,
  transform: 'rotate(45deg)',
  top: messagetype === 'success' ? '43%' : '49%',
  left: '41%',
}));

const Shadow = styled(Box)<ShadowProps>(({ theme, messagetype }) => ({
  position: 'absolute',
  width: '21%',
  height: '3%',
  opacity: theme.palette.mode === 'dark' ? 0.3 : 0.5,
  background: theme.palette.mode === 'dark' ? '#999' : '#777',
  left: '40%',
  top: '43%',
  borderRadius: '50%',
  zIndex: 1,
  animation: messagetype === 'success' ? 'scale 1s ease-in infinite' : 'move 3s ease-in-out infinite',
  '@keyframes scale': {
    '50%': {
      transform: 'scale(0.9)',
    },
  },
  '@keyframes move': {
    '0%': {
      left: '25%',
    },
    '50%': {
      left: '60%',
    },
    '100%': {
      left: '25%',
    },
  },
}));

const Message = styled(Box)(({ theme }) => ({
  textAlign: 'center',
  marginTop: '20px',
}));

const ActionButton = styled(Button)<ActionButtonProps>(({ theme, messagetype }) => ({
  background: theme.palette.mode === 'dark' ? '#2f2f2f' : '#fcfcfc',
  width: '50%',
  height: '15%',
  borderRadius: '20px',
  marginTop: '20px',
  outline: 0,
  border: 'none',
  boxShadow: theme.palette.mode === 'dark' 
    ? '0 4px 12px rgba(0, 0, 0, 0.3)' 
    : '2px 2px 10px rgba(119, 119, 119, 0.5)',
  transition: 'all 0.5s ease-in-out',
  '&:hover': {
    background: theme.palette.mode === 'dark' ? '#3f3f3f' : '#efefef',
    transform: 'scale(1.05)',
  },
}));

interface StatusMessageProps {
  type: 'success' | 'error';
  title: string;
  message: string;
  buttonText: string;
  onButtonClick: () => void;
}

const StatusMessage: React.FC<StatusMessageProps> = ({
  type,
  title,
  message,
  buttonText,
  onButtonClick,
}) => {
  return (
    <Overlay onClick={onButtonClick}>
      <Container onClick={(e) => e.stopPropagation()}>
        <MessageBox messagetype={type}>
          <Face>
            <Eye className="left" />
            <Eye className="right" />
            <Mouth messagetype={type} />
          </Face>
          <Shadow messagetype={type} />
          <Message>
            <Typography variant="h1" sx={{ 
              fontSize: '0.9em',
              fontWeight: 700,
              letterSpacing: '5px',
              color: '#fcfcfc',
              marginBottom: '10px',
            }}>
              {title}
            </Typography>
            <Typography variant="body2" sx={{ 
              fontSize: '0.5em',
              color: '#5e5e5e',
              letterSpacing: '1px',
            }}>
              {message}
            </Typography>
          </Message>
          <ActionButton 
            messagetype={type}
            onClick={onButtonClick}
            sx={{ 
              color: type === 'success' ? '#4ec07d' : '#e96075',
            }}
          >
            {buttonText}
          </ActionButton>
        </MessageBox>
      </Container>
    </Overlay>
  );
};

export default StatusMessage; 