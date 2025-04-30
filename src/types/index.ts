export interface ValidationError {
  type: 'empty' | 'format';
  message: string;
  column: string;
  rows?: number[];
} 