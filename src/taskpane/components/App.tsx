// import React, { useEffect, useState, useMemo } from "react";
// import {
//   Box, Typography, TextField, Button, Stack, MenuItem, Select,
//   FormControl, CircularProgress, IconButton, Fade, Snackbar, 
//   Alert as MuiAlert, LinearProgress, Paper, InputLabel,
//   ThemeProvider, createTheme, CssBaseline
// } from "@mui/material";
// import { 
//   CheckCircleRounded as SuccessIcon,
//   AccessTimeRounded as TimeIcon,
//   RefreshRounded as RefreshIcon,
//   AccountCircleRounded as UserIcon,
//   WbSunnyRounded as SunIcon, DarkModeRounded as MoonIcon,
//   ErrorOutlineRounded as ErrorIcon,
//   AutoAwesome as AIIcon
// } from "@mui/icons-material";

// declare const Office: any;
// const API = "https://mahfk-api-production.up.railway.app";

// // Brand Color extracted from your logo
// const BRAND_GOLD = "#C5A368"; 

// export default function App() {
//   const [mode, setMode] = useState<'light' | 'dark'>('dark');
//   const [user, setUser] = useState<any>(null);
//   const [email, setEmail] = useState("");
//   const [matters, setMatters] = useState<any[]>([]);
//   const [parsedEntry, setParsedEntry] = useState<any>(null);
//   const [matterOverride, setMatterOverride] = useState("");
//   const [hoursOverride, setHoursOverride] = useState<number | "">("");
  
//   const [isLoggingIn, setIsLoggingIn] = useState(false);
//   const [isAnalyzing, setIsAnalyzing] = useState(false);
//   const [toast, setToast] = useState({ open: false, msg: "", type: "info" as "success" | "error" | "info" | "warning" });

//   const theme = useMemo(() => createTheme({
//     palette: {
//       mode,
//       primary: {
//         main: BRAND_GOLD,
//         contrastText: mode === 'dark' ? '#000' : '#fff',
//       },
//       secondary: {
//         main: mode === 'dark' ? '#FFFFFF' : '#1E293B',
//       },
//       background: {
//         default: mode === 'dark' ? '#0F0F11' : '#F4F7F9',
//         paper: mode === 'dark' ? '#1A1A1E' : '#FFFFFF',
//       },
//       text: {
//         primary: mode === 'dark' ? '#E4E4E7' : '#1E293B',
//         secondary: mode === 'dark' ? '#A1A1AA' : '#64748B',
//       },
//     },
//     shape: { borderRadius: 12 },
//     typography: {
//       fontFamily: '"Inter", "Segoe UI", sans-serif',
//       h6: { fontWeight: 700, letterSpacing: '-0.02em' },
//       button: { textTransform: 'none', fontWeight: 700, letterSpacing: '0.02em' },
//     },
//     components: {
//       MuiButton: {
//         styleOverrides: {
//           root: { borderRadius: '8px', padding: '10px 20px' },
//           containedPrimary: {
//             background: `linear-gradient(45deg, ${BRAND_GOLD} 30%, #D4AF37 90%)`,
//             boxShadow: '0 3px 5px 2px rgba(197, 163, 104, .2)',
//           }
//         }
//       },
//       MuiPaper: {
//         styleOverrides: {
//           root: {
//             backgroundImage: 'none',
//             border: mode === 'dark' ? '1px solid rgba(255,255,255,0.05)' : '1px solid rgba(0,0,0,0.05)',
//           }
//         }
//       },
//       MuiOutlinedInput: {
//         styleOverrides: {
//           root: {
//             backgroundColor: mode === 'dark' ? 'rgba(255,255,255,0.02)' : 'rgba(0,0,0,0.01)',
//             '&:hover .MuiOutlinedInput-notchedOutline': { borderColor: BRAND_GOLD },
//           }
//         }
//       }
//     }
//   }), [mode]);

//   // Logic functions (login, loadMatters, etc. remain the same as your provided code)
//   const showToast = (msg: string, type: "success" | "error" | "info" | "warning") => setToast({ open: true, msg, type });

//   const login = async () => {
//     if (!email) { showToast("Please enter your email", "warning"); return; }
//     setIsLoggingIn(true);
//     try {
//       const res = await fetch(`${API}/auth/login`, {
//         method: "POST",
//         headers: { "Content-Type": "application/json" },
//         body: JSON.stringify({ email: email.trim() })
//       });
//       const data = await res.json();
//       if (!res.ok) { showToast(data.error || "Authentication failed", "error"); return; }
//       setUser(data);
//       localStorage.setItem("mahfk_user", JSON.stringify(data));
//       showToast("Access Granted", "success");
//     } catch (err) { showToast("Server Connection Error", "error"); }
//     finally { setIsLoggingIn(false); }
//   };

//   const loadMatters = async () => {
//     try {
//       const res = await fetch(`${API}/matters`);
//       const data = await res.json();
//       setMatters(data);
//       parseEmail();
//     } catch (err) { showToast("Could not load matters", "warning"); }
//   };

//   const parseEmail = async () => {
//     if (!user) return;
//     setIsAnalyzing(true);
//     try {
//       const item = Office.context.mailbox.item;
//       item.body.getAsync(Office.CoercionType.Text, async (result: any) => {
//         const res = await fetch(`${API}/entries/from-email`, {
//           method: "POST",
//           headers: { "Content-Type": "application/json" },
//           body: JSON.stringify({ 
//             subject: item.subject, body: (result.value || "").slice(0, 1500), 
//             sender: item.from?.emailAddress, user_id: user.id 
//           })
//         });
//         const entry = await res.json();
//         setParsedEntry(entry);
//         setMatterOverride(entry.matter_id);
//         setHoursOverride(entry.hours);
//         setIsAnalyzing(false);
//       });
//     } catch (err) { setIsAnalyzing(false); }
//   };

//   const logEntry = async () => {
//     setIsAnalyzing(true);
//     try {
//       const res = await fetch(`${API}/entries/${parsedEntry.id}`, {
//         method: "PUT",
//         headers: { "Content-Type": "application/json" },
//         body: JSON.stringify({ ...parsedEntry, matter_id: matterOverride, hours: hoursOverride })
//       });
//       if (!res.ok) throw new Error("Sync failed");
//       showToast("Entry Synchronized", "success");
//       setParsedEntry(null);
//     } catch (err: any) { showToast(err.message, "error"); }
//     finally { setIsAnalyzing(false); }
//   };

//   useEffect(() => { Office.onReady(() => { const saved = localStorage.getItem("mahfk_user"); if (saved) setUser(JSON.parse(saved)); }); }, []);
//   useEffect(() => { if (user) loadMatters(); }, [user]);

//   const activeMatter = matters.find(m => m.id === (matterOverride || parsedEntry?.matter_id));

//   return (
//     <ThemeProvider theme={theme}>
//       <CssBaseline />
//       <Box sx={{ minHeight: '100vh', display: 'flex', flexDirection: 'column', bgcolor: 'background.default' }}>
        
//         {/* Header */}
//         <Box sx={{ p: 2, display: 'flex', alignItems: 'center', justifyContent: 'space-between', borderBottom: '1px solid', borderColor: 'divider', bgcolor: 'background.paper' }}>
//           <Stack direction="row" spacing={1.5} alignItems="center">
//             <Box sx={{ bgcolor: BRAND_GOLD, p: 0.5, borderRadius: '8px', display: 'flex' }}>
//                <img src={require('../../../assets/icon-128-removebg-preview.png')} alt="Logo" width={32} height={32} />
//             </Box>
//             <Typography variant="subtitle1" sx={{ fontWeight: 900, letterSpacing: 1, color: BRAND_GOLD }}>MAHFK</Typography>
//           </Stack>
//           <IconButton onClick={() => setMode(mode === 'dark' ? 'light' : 'dark')} size="small" sx={{ border: '1px solid', borderColor: 'divider' }}>
//             {mode === 'dark' ? <SunIcon fontSize="small" /> : <MoonIcon fontSize="small" />}
//           </IconButton>
//         </Box>

//         {!user ? (
//           <Box sx={{ flex: 1, display: 'flex', alignItems: 'center', p: 4 }}>
//             <Fade in>
//               <Stack spacing={4} sx={{ width: '100%', textAlign: 'center' }}>
//                 <Box>
//                   <img src={require('../../../assets/icon-128-removebg-preview.png' )} width={80} alt="Logo" style={{ filter: mode === 'dark' ? 'drop-shadow(0 0 10px rgba(197,163,104,0.3))' : 'none' }} />
//                   <Typography variant="h5" sx={{ fontWeight: 900, mt: 2, color: BRAND_GOLD }}>MAHFK</Typography>
//                   <Typography variant="caption" sx={{ color: 'text.secondary', letterSpacing: 2 }}>LEGAL ENTERPRISE SOLUTIONS</Typography>
//                 </Box>

//                 <Stack spacing={2}>
//                   <TextField 
//                     placeholder="Corporate Email" 
//                     fullWidth 
//                     value={email}
//                     onChange={(e) => setEmail(e.target.value)}
//                     InputProps={{ startAdornment: <UserIcon sx={{ mr: 1, color: BRAND_GOLD, opacity: 0.7 }} /> }}
//                   />
//                   <Button variant="contained" fullWidth size="large" disabled={isLoggingIn} onClick={login}>
//                     {isLoggingIn ? <CircularProgress size={24} color="inherit" /> : "Verify Identity"}
//                   </Button>
//                 </Stack>
//               </Stack>
//             </Fade>
//           </Box>
//         ) : (
//           <>
//             {isAnalyzing && <LinearProgress color="primary" sx={{ height: 2 }} />}
//             <Box sx={{ flex: 1, overflowY: 'auto', p: 2 }}>
//               {!parsedEntry && !isAnalyzing ? (
//                 <Stack spacing={2} alignItems="center" sx={{ mt: 10, opacity: 0.6 }}>
//                   <TimeIcon sx={{ fontSize: 60, color: 'divider' }} />
//                   <Typography variant="body2" color="text.secondary">Select an email to begin sync</Typography>
//                   <Button variant="outlined" startIcon={<RefreshIcon />} onClick={parseEmail} color="secondary" size="small">Manual Scan</Button>
//                 </Stack>
//               ) : parsedEntry && (
//                 <Fade in>
//                   <Stack spacing={3}>
//                     <Paper sx={{ p: 2.5, position: 'relative', overflow: 'hidden' }}>
//                       <Box sx={{ position: 'absolute', top: 0, left: 0, width: '4px', height: '100%', bgcolor: BRAND_GOLD }} />
//                       <Stack direction="row" justifyContent="space-between" alignItems="flex-start">
//                         <Typography variant="caption" sx={{ color: BRAND_GOLD, fontWeight: 800, display: 'flex', alignItems: 'center', gap: 0.5 }}>
//                           <AIIcon sx={{ fontSize: 14 }} /> AI RECOGNITION
//                         </Typography>
//                       </Stack>
                      
//                       <Typography variant="h6" sx={{ mt: 1 }}>{activeMatter?.client || "General Matter"}</Typography>
//                       <Typography variant="body2" color="text.secondary" sx={{ fontStyle: 'italic', mt: 1, mb: 3, lineBreak: 'anywhere' }}>
//                         "{parsedEntry.description}"
//                       </Typography>

//                       <Stack direction="row" spacing={4}>
//                         <Box>
//                           <Typography variant="caption" color="text.secondary" display="block">Hours</Typography>
//                           <Typography variant="h5" sx={{ fontWeight: 800 }}>{hoursOverride || parsedEntry.hours}</Typography>
//                         </Box>
//                         <Box>
//                           <Typography variant="caption" color="text.secondary" display="block">Est. Value</Typography>
//                           <Typography variant="h5" sx={{ fontWeight: 800, color: 'success.main' }}>${parsedEntry.fee}</Typography>
//                         </Box>
//                       </Stack>
//                     </Paper>

//                     <Stack spacing={2}>
//                       <Typography variant="caption" sx={{ fontWeight: 800, color: 'text.secondary', ml: 1 }}>RECONCILIATION</Typography>
//                       <FormControl fullWidth size="small">
//                         <InputLabel>Client Matter</InputLabel>
//                         <Select label="Client Matter" value={matterOverride} onChange={(e) => setMatterOverride(e.target.value)}>
//                           {matters.map(m => <MenuItem key={m.id} value={m.id}>{m.client}</MenuItem>)}
//                         </Select>
//                       </FormControl>
//                       <TextField 
//                         label="Time Adjustment" type="number" fullWidth size="small"
//                         value={hoursOverride} onChange={(e) => setHoursOverride(Number(e.target.value))}
//                       />
//                     </Stack>
//                   </Stack>
//                 </Fade>
//               )}
//             </Box>

//             <Box sx={{ p: 2, bgcolor: 'background.paper', borderTop: '1px solid', borderColor: 'divider' }}>
//               {parsedEntry && (
//                 <Button variant="contained" fullWidth size="large" onClick={logEntry} disabled={isAnalyzing}>
//                   {isAnalyzing ? <CircularProgress size={24} color="inherit" /> : "Confirm & Sync Entry"}
//                 </Button>
//               )}
//             </Box>
//           </>
//         )}

//         <Snackbar open={toast.open} autoHideDuration={4000} onClose={() => setToast({ ...toast, open: false })}>
//           <MuiAlert severity={toast.type} variant="filled" sx={{ width: '100%', borderRadius: '8px' }}>
//             {toast.msg}
//           </MuiAlert>
//         </Snackbar>
//       </Box>
//     </ThemeProvider>
//   );
// }
import React, { useState, useMemo, useEffect } from "react";
import { ThemeProvider, createTheme, CssBaseline, Box, Snackbar, Alert as MuiAlert } from "@mui/material";
import Login from "./Login";
import MainUI from "./MainUI";

// Brand Color from your logo
const BRAND_GOLD = "#C5A368"; 

export default function App() {
  const [mode, setMode] = useState<'light' | 'dark'>('dark');
  const [user, setUser] = useState<any>(null);
  const [toast, setToast] = useState({ open: false, msg: "", type: "info" as any });

  const showToast = (msg: string, type: "success" | "error" | "info" | "warning") => {
    setToast({ open: true, msg, type });
  };

  // Theme configuration (Preserved from your logo style)
  const theme = useMemo(() => createTheme({
    palette: {
      mode,
      primary: { main: BRAND_GOLD, contrastText: mode === 'dark' ? '#000' : '#fff' },
      background: { default: mode === 'dark' ? '#0F0F11' : '#F4F7F9', paper: mode === 'dark' ? '#1A1A1E' : '#FFFFFF' },
    },
    shape: { borderRadius: 12 },
    typography: { fontFamily: '"Inter", sans-serif' },
    components: {
      MuiButton: { styleOverrides: { root: { borderRadius: '8px' } } },
    }
  }), [mode]);

  useEffect(() => {
    const saved = localStorage.getItem("mahfk_user");
    if (saved) setUser(JSON.parse(saved));
  }, []);

  return (
    <ThemeProvider theme={theme}>
      <CssBaseline />
      <Box sx={{ minHeight: '100vh', display: 'flex', flexDirection: 'column', bgcolor: 'background.default' }}>
        
        {/* Conditional Management */}
        {!user ? (
          <Login setUser={setUser} showToast={showToast} mode={mode} setMode={setMode} />
        ) : (
          <MainUI user={user} showToast={showToast} mode={mode} setMode={setMode} />
        )}

        <Snackbar 
            open={toast.open} 
            autoHideDuration={4000} 
            onClose={() => setToast({ ...toast, open: false })}
            anchorOrigin={{ vertical: 'bottom', horizontal: 'center' }}
        >
          <MuiAlert severity={toast.type} variant="filled" sx={{ width: '100%', borderRadius: '8px' }}>
            {toast.msg}
          </MuiAlert>
        </Snackbar>
      </Box>
    </ThemeProvider>
  );
}