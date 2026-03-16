// import React, { useState, useEffect } from "react";
// import { Box, Typography, TextField, Button, Stack, MenuItem, Select, FormControl, CircularProgress, Paper, LinearProgress, IconButton } from "@mui/material";
// import { RefreshRounded as RefreshIcon, AccessTimeRounded as TimeIcon, AutoAwesome as AIIcon, WbSunnyRounded as SunIcon, DarkModeRounded as MoonIcon } from "@mui/icons-material";

// declare const Office: any;
// const API = "https://mahfk-api-production.up.railway.app";
// const BRAND_GOLD = "#C5A368";

// export default function MainUI({ user, showToast, mode, setMode }: any) {
//   const [matters, setMatters] = useState<any[]>([]);
//   const [parsedEntry, setParsedEntry] = useState<any>(null);
//   const [matterOverride, setMatterOverride] = useState("");
//   const [hoursOverride, setHoursOverride] = useState<number | "">("");
//   const [loading, setLoading] = useState(false);

//   const loadData = async () => {
//     try {
//       const res = await fetch(`${API}/matters`);
//       const data = await res.json();
//       setMatters(data);
//       parseEmail();
//     } catch (e) { showToast("Failed to load matters", "warning"); }
//   };

//   const parseEmail = () => {
//     setLoading(true);
//     Office.context.mailbox.item.body.getAsync("text", async (res: any) => {
//       try {
//         const response = await fetch(`${API}/entries/from-email`, {
//           method: "POST",
//           headers: { "Content-Type": "application/json" },
//           body: JSON.stringify({ 
//             subject: Office.context.mailbox.item.subject,
//             body: res.value.slice(0, 1500),
//             user_id: user.id 
//           })
//         });
//         const entry = await response.json();
//         setParsedEntry(entry);
//         setMatterOverride(entry.matter_id);
//         setHoursOverride(entry.hours);
//       } catch (e) { showToast("AI Scan Failed", "error"); }
//       finally { setLoading(false); }
//     });
//   };

//   const syncEntry = async () => {
//     setLoading(true);
//     try {
//       await fetch(`${API}/entries/${parsedEntry.id}`, {
//         method: "PUT",
//         headers: { "Content-Type": "application/json" },
//         body: JSON.stringify({ ...parsedEntry, matter_id: matterOverride, hours: hoursOverride })
//       });
//       showToast("Synchronized Successfully", "success");
//       setParsedEntry(null);
//     } catch (e) { showToast("Sync Error", "error"); }
//     finally { setLoading(false); }
//   };

//   useEffect(() => { loadData(); }, []);

//   const activeMatter = matters.find(m => m.id === (matterOverride || parsedEntry?.matter_id));

//   return (
//     <>
//       <Box sx={{ p: 2, display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid', borderColor: 'divider' }}>
//         <Typography variant="subtitle2" sx={{ fontWeight: 800, color: BRAND_GOLD }}>MAHFK SYNC</Typography>
//         <IconButton onClick={() => setMode(mode === 'dark' ? 'light' : 'dark')} size="small">
//           {mode === 'dark' ? <SunIcon fontSize="small"/> : <MoonIcon fontSize="small"/>}
//         </IconButton>
//       </Box>

//       {loading && <LinearProgress color="primary" />}

//       <Box sx={{ flex: 1, p: 2, overflowY: 'auto' }}>
//         {!parsedEntry ? (
//           <Stack spacing={2} alignItems="center" sx={{ mt: 8, opacity: 0.5 }}>
//             <TimeIcon sx={{ fontSize: 48 }} />
//             <Typography variant="body2">Open an email to capture time</Typography>
//             <Button variant="outlined" onClick={parseEmail}>Manual Refresh</Button>
//           </Stack>
//         ) : (
//           <Stack spacing={3}>
//             <Paper sx={{ p: 2, borderLeft: `4px solid ${BRAND_GOLD}` }}>
//               <Typography variant="caption" sx={{ color: BRAND_GOLD, fontWeight: 900 }}>AI SUGGESTION</Typography>
//               <Typography variant="h6">{activeMatter?.client || "General Matter"}</Typography>
//               <Typography variant="body2" sx={{ fontStyle: 'italic', mt: 1 }}>"{parsedEntry.description}"</Typography>
//               <Stack direction="row" spacing={4} sx={{ mt: 2 }}>
//                 <Box><Typography variant="caption">Hours</Typography><Typography variant="h6">{hoursOverride || parsedEntry.hours}</Typography></Box>
//                 <Box><Typography variant="caption">Value</Typography><Typography variant="h6" color="success.main">${parsedEntry.fee}</Typography></Box>
//               </Stack>
//             </Paper>

//             <Stack spacing={2}>
//               <Typography variant="caption" sx={{ fontWeight: 800 }}>RECONCILIATION</Typography>
//               <Select fullWidth size="small" value={matterOverride} onChange={(e) => setMatterOverride(e.target.value)}>
//                 {matters.map(m => <MenuItem key={m.id} value={m.id}>{m.client}</MenuItem>)}
//               </Select>
//               <TextField label="Time Adj." type="number" fullWidth size="small" value={hoursOverride} onChange={(e) => setHoursOverride(Number(e.target.value))} />
//             </Stack>
//           </Stack>
//         )}
//       </Box>

//       <Box sx={{ p: 2, borderTop: '1px solid', borderColor: 'divider' }}>
//         {parsedEntry && <Button variant="contained" fullWidth size="large" onClick={syncEntry} disabled={loading}>Post Time Entry</Button>}
//       </Box>
//     </>
//   );
// }
import React, { useState, useEffect, useRef } from "react";
import { Box, Typography, TextField, Button, Stack, MenuItem, Select, IconButton, Paper, LinearProgress, Divider } from "@mui/material";
import { 
  RefreshRounded as RefreshIcon, 
  AccessTimeRounded as TimeIcon, 
  PlayArrowRounded as PlayIcon, 
  PauseRounded as PauseIcon, 
  RestartAltRounded as ResetIcon,
  WbSunnyRounded as SunIcon, 
  DarkModeRounded as MoonIcon 
} from "@mui/icons-material";

declare const Office: any;
const API = "https://mahfk-api-production.up.railway.app";
const BRAND_GOLD = "#C5A368";

export default function MainUI({ user, showToast, mode, setMode }: any) {
  const [matters, setMatters] = useState<any[]>([]);
  const [parsedEntry, setParsedEntry] = useState<any>(null);
  const [matterOverride, setMatterOverride] = useState("");
  const [hoursOverride, setHoursOverride] = useState<number | "">("");
  const [loading, setLoading] = useState(false);

  // --- TIMER STATE ---
  const [seconds, setSeconds] = useState(0);
  const [isRunning, setIsRunning] = useState(false);
  const timerRef = useRef<NodeJS.Timeout | null>(null);

  // Convert seconds to HH:MM:SS for the display
  const formatDisplayTime = (s: number) => {
    const h = Math.floor(s / 3600);
    const m = Math.floor((s % 3600) / 60);
    const sec = s % 60;
    return [h, m, sec].map(v => v < 10 ? "0" + v : v).join(":");
  };

  // Legal Billing Logic: Round UP to nearest 0.1 (6-minute block)
  // 1 sec -> 0.1, 6 mins -> 0.1, 7 mins -> 0.2
  const calculateLegalHours = (s: number) => {
    if (s === 0) return 0;
    const decimalHours = s / 3600;
    return Math.ceil(decimalHours * 10) / 10; 
  };

  useEffect(() => {
    if (isRunning) {
      timerRef.current = setInterval(() => {
        setSeconds(prev => {
          const next = prev + 1;
          // Auto-update the hours field as the timer runs
          setHoursOverride(calculateLegalHours(next));
          return next;
        });
      }, 1000);
    } else {
      if (timerRef.current) clearInterval(timerRef.current);
    }
    return () => { if (timerRef.current) clearInterval(timerRef.current); };
  }, [isRunning]);

  const resetTimer = () => {
    setIsRunning(false);
    setSeconds(0);
    setHoursOverride("");
  };
  // -------------------

  const loadData = async () => {
    try {
      const res = await fetch(`${API}/matters`);
      const data = await res.json();
      setMatters(data);
      parseEmail();
    } catch (e) { showToast("Failed to load matters", "warning"); }
  };

  const parseEmail = () => {
    setLoading(true);
    Office.context.mailbox.item.body.getAsync("text", async (res: any) => {
      try {
        const response = await fetch(`${API}/entries/from-email`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ 
            subject: Office.context.mailbox.item.subject,
            body: res.value.slice(0, 1500),
            user_id: user.id 
          })
        });
        const entry = await response.json();
        setParsedEntry(entry);
        setMatterOverride(entry.matter_id);
        // Only override hours if timer hasn't started
        if (seconds === 0) setHoursOverride(entry.hours);
      } catch (e) { showToast("AI Scan Failed", "error"); }
      finally { setLoading(false); }
    });
  };

  const syncEntry = async () => {
    setLoading(true);
    try {
      await fetch(`${API}/entries/${parsedEntry.id}`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ...parsedEntry, matter_id: matterOverride, hours: hoursOverride })
      });
      showToast("Synchronized Successfully", "success");
      setParsedEntry(null);
      resetTimer();
    } catch (e) { showToast("Sync Error", "error"); }
    finally { setLoading(false); }
  };

  useEffect(() => { loadData(); }, []);

  const activeMatter = matters.find(m => m.id === (matterOverride || parsedEntry?.matter_id));

  return (
    <>
      <Box sx={{ p: 2, display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid', borderColor: 'divider' }}>
        <Typography variant="subtitle2" sx={{ fontWeight: 800, color: BRAND_GOLD }}>MAHFK SYNC</Typography>
        <IconButton onClick={() => setMode(mode === 'dark' ? 'light' : 'dark')} size="small">
          {mode === 'dark' ? <SunIcon fontSize="small"/> : <MoonIcon fontSize="small"/>}
        </IconButton>
      </Box>

      {loading && <LinearProgress color="primary" />}

      {/* NEW TIMER WIDGET */}
      <Paper elevation={0} sx={{ m: 2, p: 2, bgcolor: 'action.hover', borderRadius: 2, textAlign: 'center' }}>
        <Typography variant="caption" sx={{ fontWeight: 700, opacity: 0.6, display: 'block', mb: 1 }}>
          LIVE SESSION TRACKER
        </Typography>
        <Typography variant="h3" sx={{ fontFamily: 'monospace', fontWeight: 700, mb: 1 }}>
          {formatDisplayTime(seconds)}
        </Typography>
        <Stack direction="row" spacing={1} justifyContent="center">
          {!isRunning ? (
            <Button startIcon={<PlayIcon />} variant="contained" color="success" onClick={() => setIsRunning(true)} size="small">Start</Button>
          ) : (
            <Button startIcon={<PauseIcon />} variant="contained" color="warning" onClick={() => setIsRunning(false)} size="small">Pause</Button>
          )}
          <Button startIcon={<ResetIcon />} variant="outlined" onClick={resetTimer} size="small">Reset</Button>
        </Stack>
      </Paper>

      <Divider />

      <Box sx={{ flex: 1, p: 2, overflowY: 'auto' }}>
        {!parsedEntry ? (
          <Stack spacing={2} alignItems="center" sx={{ mt: 4, opacity: 0.5 }}>
            <TimeIcon sx={{ fontSize: 48 }} />
            <Typography variant="body2">Open an email or start the timer</Typography>
            <Button variant="outlined" onClick={parseEmail}>Scan Email</Button>
          </Stack>
        ) : (
          <Stack spacing={3}>
            <Paper sx={{ p: 2, borderLeft: `4px solid ${BRAND_GOLD}` }}>
              <Typography variant="caption" sx={{ color: BRAND_GOLD, fontWeight: 900 }}>AI SUGGESTION</Typography>
              <Typography variant="h6">{activeMatter?.client || "General Matter"}</Typography>
              <Typography variant="body2" sx={{ fontStyle: 'italic', mt: 1 }}>"{parsedEntry.description}"</Typography>
              <Stack direction="row" spacing={4} sx={{ mt: 2 }}>
                <Box>
                    <Typography variant="caption">Billing Units (0.1)</Typography>
                    <Typography variant="h6">{hoursOverride || "0.0"}</Typography>
                </Box>
                <Box>
                    <Typography variant="caption">Value</Typography>
                    <Typography variant="h6" color="success.main">
                        ${((Number(hoursOverride) || 0) * (user.rate || 350)).toFixed(2)}
                    </Typography>
                </Box>
              </Stack>
            </Paper>

            <Stack spacing={2}>
              <Typography variant="caption" sx={{ fontWeight: 800 }}>RECONCILIATION</Typography>
              <Select fullWidth size="small" value={matterOverride} onChange={(e) => setMatterOverride(e.target.value)}>
                {matters.map(m => <MenuItem key={m.id} value={m.id}>{m.client}</MenuItem>)}
              </Select>
              <TextField 
                label="Manual Hours Adjustment" 
                type="number" 
                inputProps={{ step: 0.1 }}
                fullWidth size="small" 
                value={hoursOverride} 
                onChange={(e) => setHoursOverride(Number(e.target.value))} 
                helperText="Calculated in 0.1 (6-min) increments"
              />
            </Stack>
          </Stack>
        )}
      </Box>

      <Box sx={{ p: 2, borderTop: '1px solid', borderColor: 'divider' }}>
        {parsedEntry && (
          <Button 
            variant="contained" 
            fullWidth 
            size="large" 
            onClick={syncEntry} 
            disabled={loading || !hoursOverride}
            sx={{ bgcolor: BRAND_GOLD, '&:hover': { bgcolor: '#b08f56' } }}
          >
            Post {hoursOverride} Hours
          </Button>
        )}
      </Box>
    </>
  );
}