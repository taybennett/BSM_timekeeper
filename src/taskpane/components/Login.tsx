import React, { useState } from "react";
import { Box, Typography, TextField, Button, Stack, CircularProgress, IconButton, Fade } from "@mui/material";
import { AccountCircleRounded as UserIcon, WbSunnyRounded as SunIcon, DarkModeRounded as MoonIcon } from "@mui/icons-material";

const API = "https://mahfk-api-production.up.railway.app";
const BRAND_GOLD = "#C5A368";

export default function Login({ setUser, showToast, mode, setMode }: any) {
  const [email, setEmail] = useState("");
  const [loading, setLoading] = useState(false);

  const handleLogin = async () => {
    if (!email) { showToast("Please enter email", "warning"); return; }
    setLoading(true);
    try {
      const res = await fetch(`${API}/auth/login`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ email: email.trim() })
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Login Failed");
      
      setUser(data);
      localStorage.setItem("mahfk_user", JSON.stringify(data));
      showToast("Access Granted", "success");
    } catch (err: any) {
      showToast(err.message, "error");
    } finally {
      setLoading(false);
    }
  };

  return (
    <Fade in>
      <Box sx={{ flex: 1, display: 'flex', flexDirection: 'column', p: 4, alignItems: 'center', justifyContent: 'center' }}>
        <Box sx={{ position: 'absolute', top: 16, right: 16 }}>
          <IconButton onClick={() => setMode(mode === 'dark' ? 'light' : 'dark')}>
            {mode === 'dark' ? <SunIcon /> : <MoonIcon />}
          </IconButton>
        </Box>

        <Stack spacing={4} sx={{ width: '100%', maxWidth: 400, textAlign: 'center' }}>
          <Box>
            <img src={require('../../../assets/icon-128-removebg-preview.png')} width={80} alt="Logo" />
            <Typography variant="h5" sx={{ fontWeight: 900, mt: 2, color: BRAND_GOLD }}>MAHFK</Typography>
            <Typography variant="caption" color="text.secondary">LEGAL ENTERPRISE SOLUTIONS</Typography>
          </Box>

          <Stack spacing={2}>
            <TextField 
              placeholder="Corporate Email" fullWidth
              value={email} onChange={(e) => setEmail(e.target.value)}
              InputProps={{ startAdornment: <UserIcon sx={{ mr: 1, color: BRAND_GOLD }} /> }}
            />
            <Button variant="contained" fullWidth size="large" onClick={handleLogin} disabled={loading}>
              {loading ? <CircularProgress size={24} color="inherit" /> : "Verify Identity"}
            </Button>
          </Stack>
        </Stack>
      </Box>
    </Fade>
  );
}