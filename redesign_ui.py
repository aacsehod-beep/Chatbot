#!/usr/bin/env python3
"""AuroMate UI Redesign – full CSS overhaul + targeted HTML/JS patches"""
import re, sys

PATH = r'C:\Users\Administrator\Desktop\AU\My\university_chatbot\templates\index.html'

with open(PATH, encoding='utf-8') as f:
    c = f.read()

# ═══════════════════════════════════════════════════════════════════════════════
# 1. Inject Google Fonts
# ═══════════════════════════════════════════════════════════════════════════════
FONTS = (
    '    <link rel="preconnect" href="https://fonts.googleapis.com">\n'
    '    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>\n'
    '    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700'
    '&family=Plus+Jakarta+Sans:wght@600;700;800&display=swap" rel="stylesheet">'
)
c = c.replace(
    '<title>AuroMate - University Chatbot</title>',
    '<title>AuroMate - University Chatbot</title>\n' + FONTS
)

# ═══════════════════════════════════════════════════════════════════════════════
# 2. Replace entire <style> block
# ═══════════════════════════════════════════════════════════════════════════════
NEW_CSS = """
        /* ═══════════════════════════════════════════════
           AUROMATE  —  Modern UI Redesign
           ═══════════════════════════════════════════════ */

        /* ── CSS Custom Properties (design tokens) ── */
        :root {
            --primary:    #1abc9c;
            --primary-d:  #16a085;
            --primary-60: rgba(26,188,156,.6);
            --primary-20: rgba(26,188,156,.20);
            --primary-10: rgba(26,188,156,.10);
            --primary-06: rgba(26,188,156,.06);

            --bg:         #f0fbf8;
            --surface:    #ffffff;
            --sidebar-bg: #ffffff;
            --chat-bg:    #f4fafd;

            --text:       #1a2638;
            --text-2:     #4a5568;
            --text-muted: #8a9bb0;
            --text-inv:   #ffffff;

            --border:     #ddeef5;
            --border-lt:  #eef6fb;

            --sh-sm: 0 1px 4px rgba(0,0,0,.07);
            --sh-md: 0 4px 18px rgba(26,188,156,.12), 0 2px 6px rgba(0,0,0,.06);
            --sh-lg: 0 8px 32px rgba(26,188,156,.16), 0 4px 12px rgba(0,0,0,.08);
            --sh-xl: 0 20px 60px rgba(0,0,0,.15);

            --r-sm: 8px;  --r-md: 14px; --r-lg: 20px;
            --r-xl: 28px; --r-f: 9999px;

            --t:  .2s cubic-bezier(.4,0,.2,1);
            --ts: .4s cubic-bezier(.4,0,.2,1);

            --ff:   'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
            --ff-h: 'Plus Jakarta Sans', 'Inter', sans-serif;
        }

        [data-theme="dark"] {
            --bg:         #090f1a;
            --surface:    #101b2e;
            --sidebar-bg: #0c1625;
            --chat-bg:    #0d1829;

            --text:       #e4f0ff;
            --text-2:     #8ba4c0;
            --text-muted: #4d6880;

            --border:     #1b3049;
            --border-lt:  #142338;

            --sh-md: 0 4px 16px rgba(0,0,0,.45);
            --sh-lg: 0 8px 32px rgba(0,0,0,.55);

            --primary-20: rgba(26,188,156,.14);
            --primary-10: rgba(26,188,156,.08);
            --primary-06: rgba(26,188,156,.04);
        }

        /* ── Reset & Base ── */
        *, *::before, *::after { margin:0; padding:0; box-sizing:border-box; }
        html { height:100%; }
        body {
            font-family: var(--ff);
            background: var(--bg);
            color: var(--text);
            height: 100vh;
            overflow: hidden;
            transition: background var(--ts), color var(--ts);
            -webkit-font-smoothing: antialiased;
        }

        /* ═══════════════ WELCOME SCREEN ═══════════════ */
        .welcome-screen {
            position: fixed; inset: 0;
            background: linear-gradient(145deg, #e0f8f2 0%, #ccf0e5 40%, #c6e7f6 100%);
            display: flex; align-items: center; justify-content: center;
            z-index: 999;
            transition: opacity var(--ts), transform var(--ts);
        }
        [data-theme="dark"] .welcome-screen {
            background: linear-gradient(145deg, #07101f 0%, #0a1627 60%, #060e1c 100%);
        }
        .welcome-screen.hide { opacity:0; transform:scale(1.04); pointer-events:none; }

        .welcome-content {
            text-align: center;
            padding: 52px 48px;
            background: rgba(255,255,255,.78);
            backdrop-filter: blur(24px) saturate(180%);
            -webkit-backdrop-filter: blur(24px) saturate(180%);
            border-radius: var(--r-xl);
            border: 1px solid rgba(255,255,255,.9);
            box-shadow: var(--sh-xl), 0 0 0 1px rgba(26,188,156,.08);
            max-width: 420px; width: 90%;
            animation: wc-in .6s cubic-bezier(.34,1.56,.64,1) both;
        }
        [data-theme="dark"] .welcome-content {
            background: rgba(12,22,37,.9);
            border-color: rgba(255,255,255,.07);
        }
        @keyframes wc-in {
            from { opacity:0; transform:translateY(28px) scale(.96); }
            to   { opacity:1; transform:translateY(0) scale(1); }
        }

        .logo-circle {
            width: 96px; height: 96px; margin: 0 auto 22px;
            border-radius: 50%;
            background: linear-gradient(135deg, #1abc9c, #16a085);
            padding: 3px;
            box-shadow: 0 8px 30px rgba(26,188,156,.4);
            animation: float 3.2s ease-in-out infinite;
        }
        .logo-circle img { width:100%; height:100%; object-fit:cover; border-radius:50%; }
        @keyframes float { 0%,100%{transform:translateY(0)} 50%{transform:translateY(-7px)} }

        .welcome-title {
            font-family: var(--ff-h); font-size: 2.4rem; font-weight: 800;
            background: linear-gradient(135deg, #1abc9c, #0d9679);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            background-clip: text; letter-spacing: -.5px; margin-bottom: 8px;
        }
        .welcome-subtitle {
            color: var(--text-2); font-size: .92rem;
            margin-bottom: 32px; min-height: 1.4em;
        }
        .start-btn {
            display: inline-flex; align-items: center; gap: 8px;
            padding: 14px 38px;
            background: linear-gradient(135deg, #1abc9c, #16a085);
            color: #fff; border: none; border-radius: var(--r-f);
            font-size: 1rem; font-weight: 600; font-family: var(--ff);
            cursor: pointer; letter-spacing: .4px;
            box-shadow: 0 6px 24px rgba(26,188,156,.45);
            transition: all var(--t);
        }
        .start-btn:hover { transform:translateY(-2px); box-shadow:0 10px 32px rgba(26,188,156,.55); }
        .start-btn:active { transform:translateY(0); }

        /* ═══════════════ LAYOUT ═══════════════ */
        .main-container { display:flex; height:100vh; background:var(--bg); }

        /* ═══════════════ SIDEBAR ═══════════════ */
        .sidebar {
            width: 272px; flex-shrink: 0;
            background: var(--sidebar-bg);
            border-right: 1px solid var(--border);
            display: flex; flex-direction: column;
            overflow: hidden;
            transition: transform var(--ts), background var(--ts);
            position: relative; z-index: 200;
        }
        .sidebar::before {
            content:''; position:absolute; top:0; left:0; right:0; height:3px;
            background: linear-gradient(90deg, #1abc9c, #0d9679, #1abc9c);
        }
        .sidebar-header {
            padding: 22px 18px 18px;
            display: flex; align-items: center; gap: 12px;
            border-bottom: 1px solid var(--border-lt);
            flex-shrink: 0;
        }
        .logo-icon {
            width: 40px; height: 40px; border-radius: 11px;
            overflow: hidden; flex-shrink: 0;
            background: linear-gradient(135deg, #1abc9c, #16a085);
            padding: 2px; box-shadow: 0 3px 10px rgba(26,188,156,.3);
        }
        .logo-icon img { width:100%; height:100%; object-fit:cover; border-radius:9px; }
        .sidebar-header h2 {
            font-family: var(--ff-h); font-size: 1.15rem; font-weight: 700; flex: 1;
            background: linear-gradient(135deg, #1abc9c, #16a085);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            background-clip: text;
        }
        .clear-chat-btn {
            width: 32px; height: 32px; border-radius: 8px;
            border: 1px solid var(--border); background: var(--surface);
            color: var(--text-muted); cursor: pointer;
            display: flex; align-items: center; justify-content: center;
            transition: all var(--t); flex-shrink: 0;
        }
        .clear-chat-btn:hover {
            background: #fee2e2; border-color: #fca5a5; color: #ef4444;
            transform: scale(1.1);
        }
        [data-theme="dark"] .clear-chat-btn:hover {
            background: rgba(239,68,68,.15); border-color: rgba(239,68,68,.3);
        }

        .sidebar-body { flex:1; overflow-y:auto; padding:14px 10px; }
        .sidebar-body::-webkit-scrollbar { width:4px; }
        .sidebar-body::-webkit-scrollbar-track { background:transparent; }
        .sidebar-body::-webkit-scrollbar-thumb { background:var(--border); border-radius:4px; }

        .sidebar-section { margin-bottom: 22px; }
        .sidebar-label {
            font-size: .68rem; font-weight: 700;
            text-transform: uppercase; letter-spacing: 1px;
            color: var(--text-muted); padding: 0 8px; margin-bottom: 5px;
        }
        .service-item {
            padding: 9px 12px; border-radius: 9px; cursor: pointer;
            font-size: .87rem; font-weight: 500; color: var(--text-2);
            display: flex; align-items: center; gap: 8px;
            transition: all var(--t); position: relative; margin-bottom: 2px;
        }
        .service-item:hover { background:var(--primary-06); color:var(--text); transform:translateX(4px); }
        .service-item.active { background:var(--primary-10); color:var(--primary); font-weight:600; }
        .service-item.active::before {
            content:''; position:absolute; left:-2px; top:50%; transform:translateY(-50%);
            width:3px; height:55%; background:var(--primary); border-radius:var(--r-f);
        }

        .sidebar-footer { padding:14px 10px; border-top:1px solid var(--border-lt); flex-shrink:0; }

        /* Dark mode pill toggle */
        .theme-toggle-wrap {
            display:flex; align-items:center; justify-content:space-between;
            padding: 9px 12px; border-radius: 9px;
            cursor: pointer; transition: background var(--t);
        }
        .theme-toggle-wrap:hover { background:var(--primary-06); }
        .theme-toggle-wrap > span {
            font-size: .87rem; font-weight: 500; color: var(--text-2);
            display: flex; align-items: center; gap: 6px;
        }
        .toggle-switch {
            width: 42px; height: 23px;
            background: var(--border); border-radius: var(--r-f);
            position: relative; transition: background var(--t);
            cursor: pointer; pointer-events: none; flex-shrink: 0;
        }
        .toggle-switch::after {
            content:''; position:absolute; left:3px; top:3px;
            width:17px; height:17px; border-radius:50%;
            background: white; box-shadow:0 1px 4px rgba(0,0,0,.2);
            transition: transform var(--t);
        }
        [data-theme="dark"] .toggle-switch { background:var(--primary); }
        [data-theme="dark"] .toggle-switch::after { transform:translateX(19px); }

        /* ═══════════════ CHAT AREA ═══════════════ */
        .chat-area {
            flex:1; display:flex; flex-direction:column;
            overflow:hidden; background:var(--chat-bg);
            transition:background var(--ts); position:relative;
        }
        .chat-header {
            background: linear-gradient(135deg, #1abc9c 0%, #0d9679 100%);
            padding: 18px 26px 14px;
            display: flex; flex-direction: column; gap: 3px;
            box-shadow: 0 2px 14px rgba(26,188,156,.28);
            position: relative;
        }
        .chat-header-top { display:flex; align-items:center; gap:11px; }
        #hamburgerBtn {
            display: none;
            width: 36px; height: 36px;
            background: rgba(255,255,255,.15);
            border: 1px solid rgba(255,255,255,.25);
            border-radius: 9px; color: white; cursor: pointer;
            align-items: center; justify-content: center;
            transition: all var(--t); flex-shrink: 0;
        }
        #hamburgerBtn:hover { background:rgba(255,255,255,.25); }
        .chat-header h1 {
            font-family: var(--ff-h); font-size: 1.2rem; font-weight: 700;
            color: white; flex: 1;
        }
        .chat-header > p { color:rgba(255,255,255,.78); font-size:.81rem; padding-left:2px; }
        .ai-status {
            position:absolute; top:50%; right:26px; transform:translateY(-50%);
            display:flex; align-items:center; gap:6px;
            background:rgba(255,255,255,.14); backdrop-filter:blur(8px);
            border:1px solid rgba(255,255,255,.2); border-radius:var(--r-f); padding:5px 12px;
        }
        .status-dot { width:7px; height:7px; border-radius:50%; background:#94a3b8; }
        .status-dot.online { background:#4ade80; box-shadow:0 0 6px #4ade80; animation:pulse 2s infinite; }
        .status-dot.offline { background:#f87171; }
        @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:.55} }
        #aiStatusText { font-size:.73rem; color:rgba(255,255,255,.9); white-space:nowrap; }

        /* ═══════════════ MESSAGES ═══════════════ */
        .messages-area {
            flex:1; overflow-y:auto;
            padding: 24px 22px 16px;
            display:flex; flex-direction:column; gap:10px;
            scroll-behavior:smooth;
        }
        .messages-area::-webkit-scrollbar { width:5px; }
        .messages-area::-webkit-scrollbar-track { background:transparent; }
        .messages-area::-webkit-scrollbar-thumb { background:var(--border); border-radius:5px; }

        .message { display:flex; max-width:76%; animation:msg-in .24s cubic-bezier(.34,1.2,.64,1) both; }
        @keyframes msg-in { from{opacity:0;transform:translateY(9px) scale(.97)} to{opacity:1;transform:translateY(0) scale(1)} }
        .message.bot  { align-self:flex-start; }
        .message.user { align-self:flex-end; flex-direction:row-reverse; }

        .message-bubble { padding:12px 16px; border-radius:var(--r-lg); line-height:1.62; font-size:.9rem; position:relative; }
        .message.bot .message-bubble {
            background:var(--surface); color:var(--text);
            border-radius:5px var(--r-lg) var(--r-lg) var(--r-lg);
            border:1px solid var(--border); box-shadow:var(--sh-sm);
        }
        .message.user .message-bubble {
            background:linear-gradient(135deg,#1abc9c,#16a085); color:white;
            border-radius:var(--r-lg) 5px var(--r-lg) var(--r-lg);
            box-shadow:0 4px 16px rgba(26,188,156,.35);
        }
        .message.bot.ai .message-bubble {
            border-color:rgba(26,188,156,.22);
            box-shadow:var(--sh-sm), 0 0 0 1px rgba(26,188,156,.08);
        }

        /* Timestamps */
        .msg-time { display:block; font-size:.67rem; margin-top:5px; text-align:right; opacity:0; transition:opacity var(--t); }
        .message.user .msg-time { color:rgba(255,255,255,.6); }
        .message.bot  .msg-time { color:var(--text-muted); }
        .message:hover .msg-time { opacity:1; }

        /* Feature badges */
        .feature-badge {
            display:inline-flex; align-items:center; gap:4px;
            background:var(--primary-10); color:var(--primary);
            border-radius:var(--r-f); padding:2px 10px; font-size:.71rem; font-weight:600; margin-bottom:5px;
        }
        .auto-badge { color:var(--text-muted); font-size:.71rem; }

        /* Typing indicator */
        .typing-indicator {
            display:inline-flex; flex-direction:column; gap:7px;
            padding:12px 16px; background:var(--surface);
            border-radius:5px var(--r-lg) var(--r-lg) var(--r-lg);
            border:1px solid var(--border); box-shadow:var(--sh-sm);
        }
        .typing-label { font-size:.74rem; color:var(--text-muted); font-weight:500; }
        .typing-dots  { display:flex; gap:5px; }
        .typing-dot {
            width:8px; height:8px; border-radius:50%;
            background:var(--primary); opacity:.4;
            animation:tb 1.2s ease-in-out infinite;
        }
        .typing-dot:nth-child(2) { animation-delay:.2s; }
        .typing-dot:nth-child(3) { animation-delay:.4s; }
        @keyframes tb { 0%,80%,100%{transform:scale(.8);opacity:.4} 40%{transform:scale(1.25);opacity:1} }

        /* Suggestion chips */
        .suggestions-row { display:flex; flex-wrap:wrap; gap:7px; margin-top:4px; }
        .suggestion-chip {
            padding:7px 14px; border-radius:var(--r-f);
            border:1.5px solid var(--primary-60); background:var(--primary-06);
            color:var(--primary); font-size:.8rem; font-weight:500;
            font-family:var(--ff); cursor:pointer; transition:all var(--t);
        }
        .suggestion-chip:hover {
            background:var(--primary); color:white; border-color:var(--primary);
            transform:translateY(-2px); box-shadow:0 4px 12px rgba(26,188,156,.3);
        }

        /* ═══════════════ SCROLL BTN ═══════════════ */
        #scrollBtn {
            position:absolute; right:26px; bottom:88px;
            width:44px; height:44px; border-radius:50%;
            background:var(--primary); color:white;
            border:none; cursor:pointer;
            display:none; align-items:center; justify-content:center;
            box-shadow:0 4px 18px rgba(26,188,156,.45);
            transition:all var(--t); z-index:50;
        }
        #scrollBtn:hover { background:var(--primary-d); transform:translateY(-2px); box-shadow:0 7px 22px rgba(26,188,156,.55); }

        /* ═══════════════ INPUT AREA ═══════════════ */
        .input-area {
            padding:14px 18px; background:var(--surface);
            border-top:1px solid var(--border-lt);
            display:flex; align-items:center; gap:9px;
            transition:background var(--ts);
        }
        .input-wrapper {
            flex:1; background:var(--chat-bg);
            border:1.5px solid var(--border); border-radius:var(--r-f);
            padding:0 18px; display:flex; align-items:center; transition:all var(--t);
        }
        .input-wrapper:focus-within {
            border-color:var(--primary); background:var(--surface);
            box-shadow:0 0 0 3px rgba(26,188,156,.12);
        }
        .input-wrapper input {
            flex:1; background:transparent; border:none; outline:none;
            padding:13px 0; font-size:.9rem; font-family:var(--ff); color:var(--text);
        }
        .input-wrapper input::placeholder { color:var(--text-muted); }
        .mic-btn {
            width:44px; height:44px; border-radius:50%;
            border:1.5px solid var(--border); background:var(--surface);
            color:var(--text-muted); cursor:pointer;
            display:flex; align-items:center; justify-content:center;
            font-size:1.1rem; transition:all var(--t); flex-shrink:0;
        }
        .mic-btn:hover { border-color:var(--primary); color:var(--primary); background:var(--primary-06); }
        .mic-btn.recording { background:var(--primary); border-color:var(--primary); color:white; animation:mpulse 1s ease-in-out infinite; }
        @keyframes mpulse { 0%,100%{box-shadow:0 0 0 0 rgba(26,188,156,.4)} 50%{box-shadow:0 0 0 8px rgba(26,188,156,0)} }

        .send-btn {
            width:44px; height:44px; border-radius:50%;
            background:linear-gradient(135deg,#1abc9c,#16a085);
            border:none; color:white; cursor:pointer;
            display:flex; align-items:center; justify-content:center;
            box-shadow:0 4px 14px rgba(26,188,156,.4);
            transition:all var(--t); flex-shrink:0;
        }
        .send-btn:hover { transform:scale(1.08) translateY(-1px); box-shadow:0 6px 20px rgba(26,188,156,.5); }
        .send-btn:active { transform:scale(.95); }
        .send-btn svg { width:18px; height:18px; }

        /* ═══════════════ OVERLAY ═══════════════ */
        #sidebarOverlay {
            display:none; position:fixed; inset:0;
            background:rgba(0,0,0,.45); backdrop-filter:blur(4px);
            z-index:150; opacity:0; transition:opacity var(--t);
        }
        #sidebarOverlay.show { display:block; opacity:1; }

        /* ═══════════════ RESPONSIVE ═══════════════ */
        @media (max-width: 768px) {
            .sidebar {
                position:fixed; top:0; left:0; bottom:0;
                transform:translateX(-100%);
                box-shadow:var(--sh-lg); z-index:300;
            }
            .sidebar.open { transform:translateX(0); }
            #hamburgerBtn { display:inline-flex !important; }
            .ai-status { display:none; }
            .chat-header { padding:14px 14px 10px; }
            .messages-area { padding:16px 12px; }
            .message { max-width:92%; }
            .input-area { padding:11px 12px; }
            #scrollBtn { right:14px; bottom:78px; }
        }

        /* ═══════════════ UTILITIES ═══════════════ */
        strong { font-weight:600; }
        a { color:var(--primary); text-decoration:none; }
        a:hover { text-decoration:underline; }
        pre, code {
            font-family:'Fira Code','Courier New',monospace;
            background:var(--primary-06); border:1px solid var(--border);
            border-radius:6px; padding:2px 6px; font-size:.84em;
        }
        pre { display:block; padding:10px 14px; overflow-x:auto; }
        ::-webkit-scrollbar { width:6px; }
        ::-webkit-scrollbar-track { background:transparent; }
        ::-webkit-scrollbar-thumb { background:var(--border); border-radius:6px; }
        ::-webkit-scrollbar-thumb:hover { background:rgba(26,188,156,.5); }
"""

def _css_replace(m):
    return '<style>' + NEW_CSS + '    </style>'

c = re.sub(r'<style>[\s\S]+?</style>', _css_replace, c, count=1)

# ═══════════════════════════════════════════════════════════════════════════════
# 3. HTML patches
# ═══════════════════════════════════════════════════════════════════════════════

# 3a. Start button – add play icon
c = c.replace(
    '<button class="start-btn" onclick="startChat()">START</button>',
    '''<button class="start-btn" onclick="startChat()">
                <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="currentColor"><path d="M8 5v14l11-7z"/></svg>
                GET STARTED
            </button>'''
)

# 3b. Clear chat button – SVG trash icon instead of emoji
c = c.replace(
    '<button class="clear-chat-btn" onclick="clearChat()" title="Clear chat">🗑️</button>',
    '''<button class="clear-chat-btn" onclick="clearChat()" title="Clear chat">
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/><path d="M9 6V4h6v2"/></svg>
                </button>'''
)

# 3c. Wrap sidebar sections in sidebar-body div
OLD_SECTIONS = '''            <div class="sidebar-section">
                <div class="sidebar-label">Mode</div>
                <div class="service-item active" id="autoItem" onclick="selectService(this, null)">🤖 Auto Detect</div>
            </div>

            <div class="sidebar-section">
                <div class="sidebar-label">Academic</div>
                <div class="service-item" onclick="selectService(this, 'Student')">Student Information</div>
                <div class="service-item" onclick="selectService(this, 'Faculty')">Faculty Directory</div>
                <div class="service-item" onclick="selectService(this, 'Attendance')">Attendance Records</div>
            </div>

            <div class="sidebar-section">
                <div class="sidebar-label">Schedule</div>
                <div class="service-item" onclick="selectService(this, 'Timetable')">Class Timetable</div>
                <div class="service-item" onclick="selectService(this, 'Workload')">Workload Management</div>
            </div>

            <div class="sidebar-footer">
                <button class="theme-toggle" id="themeToggle" onclick="toggleDark()">🌙 Dark Mode</button>
            </div>'''

NEW_SECTIONS = '''            <div class="sidebar-body">
                <div class="sidebar-section">
                    <div class="sidebar-label">Mode</div>
                    <div class="service-item active" id="autoItem" onclick="selectService(this, null)">🤖 Auto Detect</div>
                </div>
                <div class="sidebar-section">
                    <div class="sidebar-label">Academic</div>
                    <div class="service-item" onclick="selectService(this, 'Student')">👤 Student Information</div>
                    <div class="service-item" onclick="selectService(this, 'Faculty')">🎓 Faculty Directory</div>
                    <div class="service-item" onclick="selectService(this, 'Attendance')">📋 Attendance Records</div>
                </div>
                <div class="sidebar-section">
                    <div class="sidebar-label">Schedule</div>
                    <div class="service-item" onclick="selectService(this, 'Timetable')">🗓️ Class Timetable</div>
                    <div class="service-item" onclick="selectService(this, 'Workload')">📊 Workload Management</div>
                </div>
            </div>

            <div class="sidebar-footer">
                <div class="theme-toggle-wrap" onclick="toggleDark()">
                    <span id="themeLabel">🌙 Dark Mode</span>
                    <div class="toggle-switch" id="themeToggle"></div>
                </div>
            </div>'''

if OLD_SECTIONS in c:
    c = c.replace(OLD_SECTIONS, NEW_SECTIONS)
    print('  ✓ Sidebar sections + dark toggle updated')
else:
    print('  ✗ WARNING: sidebar sections not found — check HTML manually')

# 3d. Chat header – wrap in chat-header-top + SVG hamburger
OLD_HEADER_INNER = (
    '                <button id="hamburgerBtn" onclick="toggleSidebar()" title="Toggle menu">☰</button>\n'
    '                <h1 id="serviceTitle">Ask Me Anything</h1>\n'
    '                <p>Ask me anything about your academics</p>'
)
NEW_HEADER_INNER = (
    '                <div class="chat-header-top">\n'
    '                    <button id="hamburgerBtn" onclick="toggleSidebar()" title="Toggle menu">\n'
    '                        <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/></svg>\n'
    '                    </button>\n'
    '                    <h1 id="serviceTitle">Ask Me Anything</h1>\n'
    '                </div>\n'
    '                <p>Ask me anything about your academics</p>'
)
if OLD_HEADER_INNER in c:
    c = c.replace(OLD_HEADER_INNER, NEW_HEADER_INNER)
    print('  ✓ Hamburger button updated')
else:
    print('  ✗ WARNING: hamburger pattern not found')

# 3e. Scroll-to-bottom button – SVG chevron
c = c.replace(
    '<button id="scrollBtn" onclick="scrollToBottom()" title="Scroll to bottom">↓</button>',
    '''<button id="scrollBtn" onclick="scrollToBottom()" title="Scroll to bottom">
                <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>
            </button>'''
)

# ═══════════════════════════════════════════════════════════════════════════════
# 4. JS patches
# ═══════════════════════════════════════════════════════════════════════════════

# 4a. toggleDark – update to reference themeLabel span, not the toggle div
OLD_TOGGLE_FN = '''        function toggleDark() {
            const btn = document.getElementById('themeToggle');
            const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
            if (isDark) {
                document.documentElement.removeAttribute('data-theme');
                btn.textContent = '\\uD83C\\uDF19 Dark Mode';
                localStorage.setItem('theme', 'light');
            } else {
                document.documentElement.setAttribute('data-theme', 'dark');
                btn.textContent = '\\u2600\\uFE0F Light Mode';
                localStorage.setItem('theme', 'dark');
            }
        }'''
NEW_TOGGLE_FN = '''        function toggleDark() {
            const label = document.getElementById('themeLabel');
            const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
            if (isDark) {
                document.documentElement.removeAttribute('data-theme');
                if (label) label.textContent = '\\uD83C\\uDF19 Dark Mode';
                localStorage.setItem('theme', 'light');
            } else {
                document.documentElement.setAttribute('data-theme', 'dark');
                if (label) label.textContent = '\\u2600\\uFE0F Light Mode';
                localStorage.setItem('theme', 'dark');
            }
        }'''
if OLD_TOGGLE_FN in c:
    c = c.replace(OLD_TOGGLE_FN, NEW_TOGGLE_FN)
    print('  ✓ toggleDark JS updated')
else:
    print('  ✗ WARNING: toggleDark function not found exactly')

# 4b. DOMContentLoaded – update themeToggle ref to themeLabel
c = c.replace(
    "            const btn = document.getElementById('themeToggle');\n"
    "                if (btn) btn.textContent = '\\u2600\\uFE0F Light Mode';",
    "            const label = document.getElementById('themeLabel');\n"
    "                if (label) label.textContent = '\\u2600\\uFE0F Light Mode';"
)

# 4c. scrollBtn – use 'flex' instead of 'block'
c = c.replace("atBottom ? 'none' : 'block'", "atBottom ? 'none' : 'flex'")

# 4d. Typing indicator – use .typing-dots class wrapper
c = c.replace(
    'typingMsg.innerHTML = `<div class="typing-indicator">${typingLabel}<div style="display:flex;gap:4px">\n'
    '                <div class="typing-dot"></div><div class="typing-dot"></div><div class="typing-dot"></div></div></div>`;',
    'typingMsg.innerHTML = `<div class="typing-indicator">${typingLabel}<div class="typing-dots"><div class="typing-dot"></div><div class="typing-dot"></div><div class="typing-dot"></div></div></div>`;'
)

# ═══════════════════════════════════════════════════════════════════════════════
# 5. Write result
# ═══════════════════════════════════════════════════════════════════════════════
with open(PATH, 'w', encoding='utf-8') as f:
    f.write(c)

print('\nDone! AuroMate UI redesigned successfully.')
