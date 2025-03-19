import { Box, Container, CssBaseline, ThemeProvider, createTheme, Typography } from '@mui/material'
import ExcelProcessor from './components/ExcelProcessor'

const theme = createTheme({
  palette: {
    mode: 'light',
    primary: {
      main: '#007AFF'
    },
    background: {
      default: '#F5F5F7',
      paper: '#FFFFFF'
    }
  },
  typography: {
    fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif',
    h4: {
      fontWeight: 600,
      letterSpacing: '-0.5px'
    }
  },
  shape: {
    borderRadius: 12
  },
  components: {
    MuiButton: {
      styleOverrides: {
        root: {
          textTransform: 'none',
          fontWeight: 500,
          padding: '8px 16px'
        }
      }
    },
    MuiContainer: {
      styleOverrides: {
        root: {
          paddingTop: '32px'
        }
      }
    }
  }
})

export default function App() {
  return (
    <ThemeProvider theme={theme}>
      <CssBaseline />
      <Container maxWidth="lg">
        <Box sx={{ my: 4 }}>
          <Typography variant="h4" component="h1" gutterBottom>
            Excel模板转换工具
          </Typography>
          <ExcelProcessor />
        </Box>
      </Container>
    </ThemeProvider>
  )
}