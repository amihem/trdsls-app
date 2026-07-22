import { useState } from 'react';
import { supabase } from './supabaseClient';

const C = { navy:"#0F1923", navyMid:"#1A3A5C", gold:"#E8C97E", red:"#E74C3C", green:"#27AE60", bg:"#F0F4F8", card:"#FFFFFF", border:"#E2EAF4", muted:"#7A8A9A" };

const inputStyle = {
  width:"100%", padding:"13px 14px", borderRadius:10, border:`1.5px solid ${C.border}`,
  fontSize:14.5, outline:"none", background:"#fff", boxSizing:"border-box", marginBottom:12
};

const btnStyle = {
  width:"100%", padding:"13px", borderRadius:10, border:"none", background:C.navy,
  color:C.gold, fontWeight:800, fontSize:14.5, cursor:"pointer", minHeight:48
};

function friendlyError(msg){
  if(!msg) return "";
  if(msg.includes("Invalid login credentials")) return "Incorrect email or password.";
  if(msg.includes("User already registered")) return "An account with this email already exists — try logging in instead.";
  if(msg.includes("Password should be at least")) return "Password must be at least 6 characters.";
  if(msg.includes("Unable to validate email")) return "Please enter a valid email address.";
  if(msg.includes("Email not confirmed")) return "Please confirm your email first (check your inbox).";
  return msg;
}

export default function Login(){
  const [mode, setMode] = useState("login"); // login | signup | forgot | otp
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [showPw, setShowPw] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");
  const [otpSent, setOtpSent] = useState(false);
  const [otpCode, setOtpCode] = useState("");

  const reset = () => { setError(""); setInfo(""); };

  const handleLogin = async (e) => {
    e.preventDefault(); reset(); setLoading(true);
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    if (error) setError(friendlyError(error.message));
    setLoading(false);
  };

  const handleSignup = async (e) => {
    e.preventDefault(); reset(); setLoading(true);
    const { error } = await supabase.auth.signUp({ email, password });
    if (error) setError(friendlyError(error.message));
    else setInfo("Signup successful! Please check your email for a confirmation link.");
    setLoading(false);
  };

  const handleForgotPassword = async (e) => {
    e.preventDefault(); reset(); setLoading(true);
    const { error } = await supabase.auth.resetPasswordForEmail(email, {
      redirectTo: window.location.origin,
    });
    if (error) setError(friendlyError(error.message));
    else setInfo("Password reset link sent to your email. Check your inbox (and spam folder).");
    setLoading(false);
  };

  const handleSendOtp = async (e) => {
    e.preventDefault(); reset(); setLoading(true);
    const { error } = await supabase.auth.signInWithOtp({ email });
    if (error) setError(friendlyError(error.message));
    else { setOtpSent(true); setInfo("A 6-digit code has been sent to your email."); }
    setLoading(false);
  };

  const handleVerifyOtp = async (e) => {
    e.preventDefault(); reset(); setLoading(true);
    const { error } = await supabase.auth.verifyOtp({ email, token: otpCode, type: "email" });
    if (error) setError(friendlyError(error.message));
    setLoading(false);
  };

  const switchMode = (m) => { setMode(m); reset(); setOtpSent(false); setPassword(""); setOtpCode(""); };

  return (
    <div style={{
      minHeight:"100vh", background:C.navy, display:"flex", alignItems:"center",
      justifyContent:"center", padding:20, boxSizing:"border-box"
    }}>
      <div style={{ width:"100%", maxWidth:380 }}>

        {/* Logo / Branding */}
        <div style={{ textAlign:"center", marginBottom:28 }}>
          <div style={{ fontSize:44, marginBottom:6 }}>🧵</div>
          <div style={{ fontSize:11, letterSpacing:3, color:C.gold, textTransform:"uppercase", fontWeight:700 }}>AMIHEM</div>
          <div style={{ fontSize:26, fontWeight:900, color:"#fff", marginTop:2 }}>Business</div>
          <div style={{ fontSize:11, color:"rgba(255,255,255,0.45)", marginTop:6, letterSpacing:0.3 }}>Sales • Inventory • Collections</div>
        </div>

        {/* Card */}
        <div style={{
          background:C.card, borderRadius:16, padding:"28px 24px",
          boxShadow:"0 20px 50px rgba(0,0,0,0.35)"
        }}>
          <div style={{ fontSize:18, fontWeight:800, color:C.navy, marginBottom:4 }}>
            {mode==="login" && "Welcome back"}
            {mode==="signup" && "Create account"}
            {mode==="forgot" && "Reset password"}
            {mode==="otp" && "Login with Email Code"}
          </div>
          <div style={{ fontSize:13, color:C.muted, marginBottom:20 }}>
            {mode==="login" && "Login to access your dashboard"}
            {mode==="signup" && "Set up your AMIHEM Business account"}
            {mode==="forgot" && "We'll email you a reset link"}
            {mode==="otp" && (otpSent ? "Enter the code sent to your email" : "We'll email you a one-time code")}
          </div>

          {/* LOGIN */}
          {mode==="login" && (
            <form onSubmit={handleLogin}>
              <input type="email" placeholder="Email" value={email} required
                onChange={e=>setEmail(e.target.value)} style={inputStyle} />
              <div style={{ position:"relative" }}>
                <input type={showPw?"text":"password"} placeholder="Password" value={password} required
                  onChange={e=>setPassword(e.target.value)} style={{...inputStyle, paddingRight:44}} />
                <span onClick={()=>setShowPw(!showPw)} style={{
                  position:"absolute", right:14, top:13, cursor:"pointer", fontSize:16, color:C.muted
                }}>{showPw ? "🙈" : "👁️"}</span>
              </div>
              <div style={{ textAlign:"right", marginBottom:14 }}>
                <span onClick={()=>switchMode("forgot")} style={{ fontSize:12.5, color:C.navyMid, cursor:"pointer", fontWeight:600 }}>
                  Forgot password?
                </span>
              </div>
              <button type="submit" disabled={loading} style={{...btnStyle, opacity:loading?0.7:1}}>
                {loading ? "Please wait…" : "Login"}
              </button>
            </form>
          )}

          {/* SIGNUP */}
          {mode==="signup" && (
            <form onSubmit={handleSignup}>
              <input type="email" placeholder="Email" value={email} required
                onChange={e=>setEmail(e.target.value)} style={inputStyle} />
              <input type={showPw?"text":"password"} placeholder="Password (min 6 characters)" value={password} required
                onChange={e=>setPassword(e.target.value)} style={inputStyle} />
              <button type="submit" disabled={loading} style={{...btnStyle, opacity:loading?0.7:1}}>
                {loading ? "Please wait…" : "Sign Up"}
              </button>
            </form>
          )}

          {/* FORGOT PASSWORD */}
          {mode==="forgot" && (
            <form onSubmit={handleForgotPassword}>
              <input type="email" placeholder="Email" value={email} required
                onChange={e=>setEmail(e.target.value)} style={inputStyle} />
              <button type="submit" disabled={loading} style={{...btnStyle, opacity:loading?0.7:1}}>
                {loading ? "Sending…" : "Send Reset Link"}
              </button>
            </form>
          )}

          {/* EMAIL OTP */}
          {mode==="otp" && !otpSent && (
            <form onSubmit={handleSendOtp}>
              <input type="email" placeholder="Email" value={email} required
                onChange={e=>setEmail(e.target.value)} style={inputStyle} />
              <button type="submit" disabled={loading} style={{...btnStyle, opacity:loading?0.7:1}}>
                {loading ? "Sending…" : "Send Code"}
              </button>
            </form>
          )}
          {mode==="otp" && otpSent && (
            <form onSubmit={handleVerifyOtp}>
              <input type="text" inputMode="numeric" placeholder="6-digit code" value={otpCode} required
                onChange={e=>setOtpCode(e.target.value)} style={{...inputStyle, textAlign:"center", letterSpacing:4, fontSize:18}} />
              <button type="submit" disabled={loading} style={{...btnStyle, opacity:loading?0.7:1}}>
                {loading ? "Verifying…" : "Verify & Login"}
              </button>
            </form>
          )}

          {/* Messages */}
          {error && <div style={{ marginTop:14, padding:"10px 12px", background:"#FDEDEC", color:C.red, borderRadius:8, fontSize:13, fontWeight:600 }}>{error}</div>}
          {info && <div style={{ marginTop:14, padding:"10px 12px", background:"#EAFAF1", color:C.green, borderRadius:8, fontSize:13, fontWeight:600 }}>{info}</div>}

          {/* Switch links */}
          <div style={{ marginTop:20, textAlign:"center", fontSize:13, color:C.muted }}>
            {mode==="login" && (
              <>
                <div style={{ marginBottom:8 }}>
                  No account?{" "}
                  <span onClick={()=>switchMode("signup")} style={{ color:C.navyMid, fontWeight:700, cursor:"pointer" }}>Sign Up</span>
                </div>
                <div>
                  Or{" "}
                  <span onClick={()=>switchMode("otp")} style={{ color:C.navyMid, fontWeight:700, cursor:"pointer" }}>login with email code</span>
                </div>
              </>
            )}
            {mode==="signup" && (
              <div>
                Already have an account?{" "}
                <span onClick={()=>switchMode("login")} style={{ color:C.navyMid, fontWeight:700, cursor:"pointer" }}>Login</span>
              </div>
            )}
            {(mode==="forgot" || mode==="otp") && (
              <div>
                <span onClick={()=>switchMode("login")} style={{ color:C.navyMid, fontWeight:700, cursor:"pointer" }}>← Back to Login</span>
              </div>
            )}
          </div>
        </div>

        <div style={{ textAlign:"center", marginTop:20, fontSize:11.5, color:"rgba(255,255,255,0.4)" }}>
          Your data is securely stored and synced across devices.
        </div>
      </div>
    </div>
  );
}
