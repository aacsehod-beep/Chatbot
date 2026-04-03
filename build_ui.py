with open('logo_b64.txt', encoding='utf-8') as f:
    LOGO = f.read().strip()

HTML = '''<!DOCTYPE html>
<html lang="en" data-theme="light">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AuroMate - College Assistant</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Inter:ital,opsz,wght@0,14..32,300..700;1,14..32,300..700&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary: #1abc9c;
            --primary-dark: #16a085;
            --primary-glow: rgba(26,188,156,0.18);
            --bg: #f0f4f8;
            --surface: #ffffff;
            --sidebar-bg: #ffffff;
            --border: #e2e8f0;
            --text: #1a202c;
            --text-muted: #718096;
            --user-bubble-bg: linear-gradient(135deg,#1abc9c,#16a085);
            --bot-bubble-bg: #ffffff;
            --bot-bubble-border: #e2e8f0;
            --input-bg: #f7fafc;
            --shadow-sm: 0 1px 3px rgba(0,0,0,0.08);
            --shadow-md: 0 4px 12px rgba(0,0,0,0.08);
            --shadow-lg: 0 10px 30px rgba(0,0,0,0.12);
        }
        [data-theme="dark"] {
            --bg: #0f172a;
            --surface: #1e293b;
            --sidebar-bg: #1a2744;
            --border: #2d3f5e;
            --text: #e2e8f0;
            --text-muted: #94a3b8;
            --bot-bubble-bg: #243352;
            --bot-bubble-border: #2d3f5e;
            --input-bg: #0f172a;
            --shadow-sm: 0 1px 3px rgba(0,0,0,0.3);
            --shadow-md: 0 4px 12px rgba(0,0,0,0.3);
        }
        * { margin:0; padding:0; box-sizing:border-box; }
        body {
            font-family: \'Inter\', \'Segoe UI\', sans-serif;
            background: var(--bg);
            height: 100dvh;
            overflow: hidden;
            color: var(--text);
            transition: background 0.25s, color 0.25s;
        }

        /* ===== WELCOME SCREEN ===== */
        .welcome-screen {
            position: fixed; inset: 0;
            background: linear-gradient(145deg, #1abc9c 0%, #0d7d66 100%);
            display: flex; align-items: center; justify-content: center;
            z-index: 1000;
        }
        .welcome-screen.hide { animation: wsFadeOut 0.7s ease forwards; }
        @keyframes wsFadeOut { to { opacity:0; pointer-events:none; } }
        .welcome-content { text-align: center; color: white; padding: 2rem; }
        .logo-circle {
            width: 130px; height: 130px;
            background: white; border-radius: 28px;
            margin: 0 auto 28px; overflow: hidden;
            box-shadow: 0 20px 60px rgba(0,0,0,0.22);
            animation: wsLogoIn 0.9s cubic-bezier(0.34,1.56,0.64,1);
        }
        .logo-circle img { width:100%; height:100%; object-fit:contain; }
        @keyframes wsLogoIn {
            from { transform: scale(0) rotate(-15deg); opacity:0; }
            to   { transform: scale(1) rotate(0); opacity:1; }
        }
        .welcome-title {
            font-size: 3rem; font-weight: 800; letter-spacing: -1px;
            animation: wsSlideUp 0.6s ease 0.2s both;
        }
        .welcome-subtitle {
            font-size: 1rem; opacity: 0.85; margin: 8px 0 36px;
            animation: wsSlideUp 0.6s ease 0.35s both;
        }
        @keyframes wsSlideUp {
            from { opacity:0; transform: translateY(20px); }
        }
        .start-btn {
            padding: 14px 52px;
            background: white; color: #16a085;
            border: none; border-radius: 50px;
            font-size: 1rem; font-weight: 700; letter-spacing: 2px;
            cursor: pointer; text-transform: uppercase;
            box-shadow: 0 15px 40px rgba(0,0,0,0.2);
            transition: transform 0.25s, box-shadow 0.25s;
            animation: wsSlideUp 0.6s ease 0.5s both;
        }
        .start-btn:hover { transform: translateY(-3px); box-shadow: 0 22px 50px rgba(0,0,0,0.28); }

        /* ===== LAYOUT ===== */
        .main-container { display:flex; height:100dvh; background:var(--bg); }

        /* ===== SIDEBAR OVERLAY ===== */
        #sidebarOverlay {
            display: none; position: fixed; inset: 0;
            background: rgba(0,0,0,0.45); z-index: 40;
            backdrop-filter: blur(2px);
        }
        #sidebarOverlay.visible { display: block; }

        /* ===== SIDEBAR ===== */
        .sidebar {
            width: 265px;
            background: var(--sidebar-bg);
            border-right: 1px solid var(--border);
            display: flex; flex-direction: column;
            flex-shrink: 0; z-index: 50;
            transition: transform 0.28s ease, background 0.25s;
        }
        .sidebar-topbar {
            display: flex; align-items: center;
            padding: 16px 14px 14px;
            border-bottom: 1px solid var(--border);
            gap: 10px;
        }
        .logo-icon {
            width: 40px; height: 40px; border-radius: 10px;
            overflow: hidden; flex-shrink: 0;
            border: 1px solid var(--border);
            background: white;
        }
        .logo-icon img { width:100%; height:100%; object-fit:contain; }
        .sidebar-brand { flex: 1; min-width: 0; }
        .sidebar-brand h2 { font-size: 14.5px; font-weight: 700; color: var(--primary); }
        .sidebar-brand p { font-size: 10.5px; color: var(--text-muted); margin-top:1px; }
        .clear-btn {
            width: 32px; height: 32px; flex-shrink: 0;
            background: var(--surface); border: 1px solid var(--border);
            border-radius: 8px; cursor: pointer;
            display: flex; align-items: center; justify-content: center;
            color: var(--text-muted); transition: all 0.2s;
        }
        .clear-btn:hover { background:#fee2e2; border-color:#fca5a5; color:#ef4444; }

        .sidebar-body { flex:1; overflow-y:auto; padding: 12px 12px; }
        .sidebar-body::-webkit-scrollbar { width: 4px; }
        .sidebar-body::-webkit-scrollbar-thumb { background: var(--border); border-radius: 2px; }

        .sidebar-label {
            font-size: 10px; font-weight: 700;
            text-transform: uppercase; letter-spacing: 1.2px;
            color: var(--text-muted); padding: 14px 6px 6px;
        }
        .sidebar-label:first-child { padding-top: 4px; }
        .service-item {
            display: flex; align-items: center; gap: 10px;
            padding: 9px 11px; margin-bottom: 2px;
            border-radius: 9px; cursor: pointer;
            font-size: 13px; font-weight: 500; color: var(--text);
            transition: background 0.18s, color 0.18s;
        }
        .service-item:hover { background: var(--primary-glow); color: var(--primary); }
        .service-item.active { background: var(--primary); color: white; }
        .s-icon { font-size: 15px; flex-shrink: 0; }
        .ai-pill {
            margin-left: auto; font-size: 9px; font-weight: 700;
            padding: 2px 7px; border-radius: 20px;
            background: var(--primary); color: white; letter-spacing: 0.5px;
        }
        .service-item.active .ai-pill { background: rgba(255,255,255,0.28); }

        /* ===== CHAT AREA ===== */
        .chat-area {
            flex: 1; display: flex; flex-direction: column;
            background: var(--surface); min-width: 0;
            transition: background 0.25s;
        }

        /* ===== HEADER ===== */
        .chat-header {
            background: linear-gradient(135deg, #1abc9c 0%, #16a085 100%);
            padding: 13px 18px;
            display: flex; align-items: center; gap: 10px;
            box-shadow: 0 2px 12px rgba(26,188,156,0.22);
            flex-shrink: 0;
        }
        #hamburgerBtn {
            display: none; width: 36px; height: 36px;
            background: rgba(255,255,255,0.15); border: none;
            border-radius: 8px; cursor: pointer; color: white;
            align-items: center; justify-content: center;
            flex-shrink: 0; transition: background 0.18s;
        }
        #hamburgerBtn:hover { background: rgba(255,255,255,0.25); }
        .header-info { flex: 1; min-width: 0; }
        .header-info h1 {
            font-size: 15px; font-weight: 700; color: white;
            white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
        }
        .header-info p { font-size: 11px; color: rgba(255,255,255,0.78); margin-top: 1px; }
        .header-right { display: flex; align-items: center; gap: 8px; flex-shrink: 0; }
        .ai-status {
            display: flex; align-items: center; gap: 5px;
            background: rgba(255,255,255,0.15);
            padding: 5px 10px; border-radius: 20px;
        }
        .status-dot { width: 7px; height: 7px; border-radius: 50%; background: #94a3b8; flex-shrink:0; }
        .status-dot.online  { background: #4ade80; animation: dotPulse 2s infinite; }
        .status-dot.offline { background: #f87171; }
        @keyframes dotPulse { 0%,100%{opacity:1;} 50%{opacity:0.35;} }
        #aiStatusText { font-size: 11px; color: rgba(255,255,255,0.9); font-weight: 500; white-space:nowrap; }
        .header-icon-btn {
            width: 34px; height: 34px;
            background: rgba(255,255,255,0.15); border: none;
            border-radius: 8px; cursor: pointer; color: white;
            display: flex; align-items: center; justify-content: center;
            flex-shrink: 0; transition: background 0.18s;
        }
        .header-icon-btn:hover { background: rgba(255,255,255,0.26); }

        /* ===== MESSAGES ===== */
        .messages-container {
            flex: 1; position: relative; overflow: hidden;
            display: flex; flex-direction: column;
        }
        .messages-area {
            flex: 1; overflow-y: auto; overflow-x: hidden;
            padding: 18px 18px 12px; scroll-behavior: smooth;
            display: flex; flex-direction: column; gap: 10px;
        }
        .messages-area::-webkit-scrollbar { width: 5px; }
        .messages-area::-webkit-scrollbar-thumb { background: var(--border); border-radius: 3px; }

        .message { display: flex; animation: msgPop 0.3s cubic-bezier(0.34,1.4,0.64,1); }
        @keyframes msgPop { from { opacity:0; transform: scale(0.88); } }
        .message.user { justify-content: flex-end; }
        .message-wrap { display: flex; flex-direction: column; max-width: 72%; }
        .message.user .message-wrap { align-items: flex-end; }
        .message.bot  .message-wrap { align-items: flex-start; }

        .feature-badge {
            display: inline-flex; align-items: center; gap: 4px;
            font-size: 10px; font-weight: 600;
            padding: 3px 9px; border-radius: 20px;
            background: var(--primary-glow);
            color: var(--primary);
            border: 1px solid rgba(26,188,156,0.25);
            margin-bottom: 5px;
        }
        .message-bubble {
            padding: 10px 14px; border-radius: 16px;
            font-size: 13.5px; line-height: 1.58;
            word-break: break-word;
            box-shadow: var(--shadow-sm);
        }
        .message.bot  .message-bubble {
            background: var(--bot-bubble-bg);
            color: var(--text);
            border: 1px solid var(--bot-bubble-border);
            border-radius: 4px 16px 16px 16px;
        }
        .message.user .message-bubble {
            background: var(--user-bubble-bg);
            color: white;
            border-radius: 16px 4px 16px 16px;
        }
        .msg-time {
            font-size: 10px; color: var(--text-muted);
            margin-top: 3px; opacity: 0; transition: opacity 0.2s;
        }
        .message:hover .msg-time { opacity: 1; }

        /* Typing */
        .typing-indicator {
            display: flex; align-items: center; gap: 5px;
            padding: 10px 14px;
            background: var(--bot-bubble-bg);
            border: 1px solid var(--bot-bubble-border);
            border-radius: 4px 16px 16px 16px;
            box-shadow: var(--shadow-sm); width: fit-content;
        }
        .typing-dot {
            width: 7px; height: 7px; border-radius: 50%;
            background: var(--primary);
            animation: typBounce 1.2s ease infinite;
        }
        .typing-dot:nth-child(2) { animation-delay: 0.2s; }
        .typing-dot:nth-child(3) { animation-delay: 0.4s; }
        @keyframes typBounce { 0%,60%,100%{transform:translateY(0);opacity:0.4;} 30%{transform:translateY(-7px);opacity:1;} }
        .typing-label { font-size: 10px; color: var(--text-muted); margin-top: 4px; padding-left: 2px; }

        /* Scroll btn */
        #scrollBtn {
            position: absolute; bottom: 12px; right: 18px;
            width: 36px; height: 36px;
            background: var(--primary); border: none; border-radius: 50%;
            color: white; cursor: pointer;
            display: none; align-items: center; justify-content: center;
            box-shadow: 0 4px 14px rgba(26,188,156,0.4); z-index: 5;
            transition: box-shadow 0.2s, transform 0.2s;
        }
        #scrollBtn:hover { box-shadow: 0 6px 20px rgba(26,188,156,0.5); transform: translateY(-1px); }

        /* Suggestions */
        .suggestions-row {
            display: flex; flex-wrap: wrap; gap: 6px;
            padding: 6px 18px 10px;
        }
        .suggestion-chip {
            padding: 5px 13px;
            background: var(--surface); border: 1px solid var(--border);
            border-radius: 20px; font-size: 12px; font-weight: 500;
            color: var(--primary); cursor: pointer;
            transition: all 0.18s;
        }
        .suggestion-chip:hover { background: var(--primary); color: white; border-color: var(--primary); }

        /* ===== INPUT AREA ===== */
        .input-area {
            padding: 10px 14px;
            background: var(--surface);
            border-top: 1px solid var(--border);
            display: flex; align-items: center; gap: 8px;
            flex-shrink: 0;
        }
        .input-wrapper {
            flex: 1; display: flex; align-items: center;
            background: var(--input-bg);
            border: 1.5px solid var(--border);
            border-radius: 24px; padding: 4px 6px 4px 16px;
            gap: 4px; transition: border-color 0.2s, box-shadow 0.2s;
        }
        .input-wrapper:focus-within {
            border-color: var(--primary);
            box-shadow: 0 0 0 3px rgba(26,188,156,0.12);
        }
        .input-wrapper input {
            flex: 1; border: none; background: transparent;
            font-size: 13.5px; outline: none;
            color: var(--text); padding: 6px 0;
            font-family: inherit;
        }
        .input-wrapper input::placeholder { color: var(--text-muted); }
        .mic-btn {
            width: 32px; height: 32px; border: none;
            border-radius: 50%; background: transparent;
            color: var(--text-muted); cursor: pointer;
            display: flex; align-items: center; justify-content: center;
            flex-shrink: 0; transition: all 0.2s;
        }
        .mic-btn:hover { background: var(--primary-glow); color: var(--primary); }
        .mic-btn.recording { background: #fee2e2; color: #ef4444; animation: recPulse 1.5s infinite; }
        @keyframes recPulse {
            0%,100% { box-shadow: 0 0 0 0 rgba(239,68,68,0.4); }
            50%      { box-shadow: 0 0 0 8px rgba(239,68,68,0); }
        }
        .send-btn {
            width: 42px; height: 42px; flex-shrink: 0;
            background: linear-gradient(135deg, #1abc9c, #16a085);
            border: none; border-radius: 50%;
            color: white; cursor: pointer;
            display: flex; align-items: center; justify-content: center;
            box-shadow: 0 4px 12px rgba(26,188,156,0.35);
            transition: transform 0.2s, box-shadow 0.2s;
        }
        .send-btn:hover { transform: translateY(-2px); box-shadow: 0 6px 18px rgba(26,188,156,0.45); }
        .send-btn:active { transform: scale(0.94); }

        /* ===== MOBILE ===== */
        @media (max-width: 768px) {
            .sidebar {
                position: fixed; top:0; left:0; height:100%;
                transform: translateX(-100%);
                width: 275px;
                box-shadow: var(--shadow-lg);
            }
            .sidebar.open { transform: translateX(0); }
            #hamburgerBtn { display: flex; }
            .ai-status { display: none; }
        }
        @media (max-width: 480px) {
            .messages-area { padding: 12px 12px 10px; }
            .input-area { padding: 8px 10px; }
            .message-wrap { max-width: 85%; }
        }
    </style>
</head>
<body>

<!-- Welcome Screen -->
<div class="welcome-screen" id="welcomeScreen">
    <div class="welcome-content">
        <div class="logo-circle">
            <img src="data:image/png;base64,LOGO_PLACEHOLDER" alt="College Logo">
        </div>
        <h1 class="welcome-title">AuroMate</h1>
        <p class="welcome-subtitle">Your Smart College Assistant</p>
        <button class="start-btn" onclick="startChat()">GET STARTED</button>
    </div>
</div>

<!-- Sidebar overlay (mobile) -->
<div id="sidebarOverlay" onclick="toggleSidebar(false)"></div>

<!-- Main Layout -->
<div class="main-container" id="mainContainer" style="display:none;">

    <!-- Sidebar -->
    <div class="sidebar" id="sidebar">
        <div class="sidebar-topbar">
            <div class="logo-icon">
                <img src="data:image/png;base64,LOGO_PLACEHOLDER" alt="Logo">
            </div>
            <div class="sidebar-brand">
                <h2>AuroMate</h2>
                <p>College Assistant</p>
            </div>
            <button class="clear-btn" onclick="clearChat()" title="Clear chat">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
                    <polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/>
                    <path d="M10 11v6M14 11v6"/><path d="M9 6V4h6v2"/>
                </svg>
            </button>
        </div>
        <div class="sidebar-body">
            <div class="sidebar-label">Smart</div>
            <div class="service-item" id="autoItem" onclick="selectService(this, null)">
                <span class="s-icon">🔍</span>
                Auto Detect
                <span class="ai-pill">AI</span>
            </div>
            <div class="sidebar-label">Academic</div>
            <div class="service-item active" onclick="selectService(this, \'Student\')">
                <span class="s-icon">🎓</span> Student Information
            </div>
            <div class="service-item" onclick="selectService(this, \'Faculty\')">
                <span class="s-icon">👩\u200d🏫</span> Faculty Directory
            </div>
            <div class="service-item" onclick="selectService(this, \'Attendance\')">
                <span class="s-icon">📊</span> Attendance
            </div>
            <div class="sidebar-label">Schedule</div>
            <div class="service-item" onclick="selectService(this, \'Timetable\')">
                <span class="s-icon">🗓️</span> Timetable
            </div>
            <div class="service-item" onclick="selectService(this, \'Workload\')">
                <span class="s-icon">📋</span> Workload
            </div>
        </div>
    </div>

    <!-- Chat Area -->
    <div class="chat-area">

        <!-- Header -->
        <div class="chat-header">
            <button id="hamburgerBtn" onclick="toggleSidebar(true)" aria-label="Menu">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
                    <line x1="3" y1="6" x2="21" y2="6"/>
                    <line x1="3" y1="12" x2="21" y2="12"/>
                    <line x1="3" y1="18" x2="21" y2="18"/>
                </svg>
            </button>
            <div class="header-info">
                <h1 id="serviceTitle">Student Information</h1>
                <p>Ask me anything about your academics</p>
            </div>
            <div class="header-right">
                <div class="ai-status">
                    <div class="status-dot" id="statusDot"></div>
                    <span id="aiStatusText">Checking...</span>
                </div>
                <button class="header-icon-btn" onclick="toggleDark()" title="Toggle dark mode" aria-label="Dark mode">
                    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <circle cx="12" cy="12" r="5"/>
                        <line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/>
                        <line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/>
                        <line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/>
                        <line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/>
                    </svg>
                </button>
            </div>
        </div>

        <!-- Messages -->
        <div class="messages-container">
            <div class="messages-area" id="messagesArea">
                <div class="message bot">
                    <div class="message-wrap">
                        <div class="message-bubble">👋 Welcome! I\'m <strong>AuroMate</strong>, your smart college assistant.<br>Ask me about student info, faculty, attendance, timetable, or workload — or select <em>Auto Detect</em> and let AI figure it out!</div>
                        <div class="msg-time">Just now</div>
                    </div>
                </div>
            </div>
            <button id="scrollBtn" onclick="scrollToBottom()" aria-label="Scroll to bottom">
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
                    <polyline points="6 9 12 15 18 9"/>
                </svg>
            </button>
        </div>

        <!-- Suggestions row -->
        <div class="suggestions-row" id="suggestionsRow" style="display:none;"></div>

        <!-- Input -->
        <div class="input-area">
            <div class="input-wrapper">
                <input type="text" id="userInput" placeholder="Ask anything..." autocomplete="off"
                       onkeydown="if(event.key===\'Enter\'){sendMessage();}">
                <button class="mic-btn" id="micBtn" onclick="toggleVoice()" title="Voice input" aria-label="Voice">
                    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <rect x="9" y="2" width="6" height="12" rx="3"/>
                        <path d="M5 10c0 4 2.686 7 7 7s7-3 7-7"/>
                        <line x1="12" y1="19" x2="12" y2="23"/>
                        <line x1="8" y1="23" x2="16" y2="23"/>
                    </svg>
                </button>
            </div>
            <button class="send-btn" onclick="sendMessage()" aria-label="Send">
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
                    <line x1="22" y1="2" x2="11" y2="13"/>
                    <polygon points="22 2 15 22 11 13 2 9 22 2"/>
                </svg>
            </button>
        </div>

    </div><!-- end chat-area -->
</div><!-- end main-container -->

<script>
    // ========================================
    // STATE
    // ========================================
    let selectedService = \'Student\';
    let convId = null;
    let isRecording = false;
    let recog = null;

    const SERVICE_NAMES = {
        \'Student\': \'Student Information\',
        \'Faculty\': \'Faculty Directory\',
        \'Attendance\': \'Attendance Records\',
        \'Timetable\': \'Class Timetable\',
        \'Workload\': \'Workload Management\',
        null: \'Auto Detect\'
    };
    const MODULE_ICONS = { Student:\'🎓\', Faculty:\'👩‍🏫\', Attendance:\'📊\', Timetable:\'🗓️\', Workload:\'📋\' };
    const MODULE_KEYWORDS = {
        Student: [\'student\',\'roll\',\'id\',\'gpa\',\'grade\',\'cgpa\',\'marks\',\'result\',\'admission\',\'exam\'],
        Faculty: [\'faculty\',\'professor\',\'teacher\',\'department\',\'staff\',\'lecturer\',\'hod\'],
        Attendance: [\'attendance\',\'absent\',\'present\',\'leave\',\'percentage\',\'bunk\'],
        Timetable: [\'timetable\',\'schedule\',\'class\',\'period\',\'timing\',\'lecture\',\'slot\',\'when\'],
        Workload: [\'workload\',\'load\',\'subject\',\'credit\',\'hour\',\'assign\',\'lab\']
    };

    // ========================================
    // HELPERS
    // ========================================
    function getTime() {
        return new Date().toLocaleTimeString([], {hour:\'2-digit\', minute:\'2-digit\'});
    }
    function guessModule(text) {
        const lower = text.toLowerCase();
        for (const [mod, kws] of Object.entries(MODULE_KEYWORDS)) {
            if (kws.some(k => lower.includes(k))) return mod;
        }
        return null;
    }
    function escHtml(s) {
        return s.replace(/&/g,\'&amp;\').replace(/</g,\'&lt;\').replace(/>/g,\'&gt;\');
    }

    // ========================================
    // THEME (load saved on startup)
    // ========================================
    (function() {
        const t = localStorage.getItem(\'auromate_theme\');
        if (t) document.documentElement.setAttribute(\'data-theme\', t);
    })();

    function toggleDark() {
        const cur = document.documentElement.getAttribute(\'data-theme\');
        const next = cur === \'dark\' ? \'light\' : \'dark\';
        document.documentElement.setAttribute(\'data-theme\', next);
        localStorage.setItem(\'auromate_theme\', next);
    }

    // ========================================
    // WELCOME
    // ========================================
    function startChat() {
        const ws = document.getElementById(\'welcomeScreen\');
        const mc = document.getElementById(\'mainContainer\');
        ws.classList.add(\'hide\');
        setTimeout(() => {
            ws.style.display = \'none\';
            mc.style.display = \'flex\';
            document.getElementById(\'userInput\').focus();
        }, 700);
        loadAiStatus();
        ensureConversation();
    }

    // ========================================
    // AI STATUS
    // ========================================
    async function loadAiStatus() {
        try {
            const r = await fetch(\'/ai-status\');
            const d = await r.json();
            const dot = document.getElementById(\'statusDot\');
            const txt = document.getElementById(\'aiStatusText\');
            if (d.configured) {
                dot.className = \'status-dot online\';
                txt.textContent = d.model || \'AI Ready\';
            } else {
                dot.className = \'status-dot offline\';
                txt.textContent = \'Rule-based\';
            }
        } catch {
            document.getElementById(\'statusDot\').className = \'status-dot offline\';
            document.getElementById(\'aiStatusText\').textContent = \'Offline\';
        }
    }

    // ========================================
    // CONVERSATION
    // ========================================
    async function ensureConversation() {
        if (convId) return;
        try {
            const r = await fetch(\'/new-conversation\', { method: \'POST\' });
            const d = await r.json();
            convId = d.convId;
        } catch { convId = null; }
    }

    // ========================================
    // SIDEBAR
    // ========================================
    function selectService(el, service) {
        document.querySelectorAll(\'.service-item\').forEach(i => i.classList.remove(\'active\'));
        el.classList.add(\'active\');
        selectedService = service;
        document.getElementById(\'serviceTitle\').textContent = SERVICE_NAMES[service] || \'AuroMate\';
        toggleSidebar(false);
    }
    function toggleSidebar(open) {
        document.getElementById(\'sidebar\').classList.toggle(\'open\', open);
        document.getElementById(\'sidebarOverlay\').classList.toggle(\'visible\', open);
    }

    // ========================================
    // SCROLL
    // ========================================
    function scrollToBottom() {
        const a = document.getElementById(\'messagesArea\');
        a.scrollTo({ top: a.scrollHeight, behavior: \'smooth\' });
    }
    document.addEventListener(\'DOMContentLoaded\', () => {
        document.getElementById(\'messagesArea\').addEventListener(\'scroll\', function() {
            const fromBot = this.scrollHeight - this.scrollTop - this.clientHeight;
            document.getElementById(\'scrollBtn\').style.display = fromBot > 160 ? \'flex\' : \'none\';
        });
    });

    // ========================================
    // CLEAR CHAT
    // ========================================
    function clearChat() {
        const a = document.getElementById(\'messagesArea\');
        a.innerHTML = \'<div class="message bot"><div class="message-wrap"><div class="message-bubble">Chat cleared. How can I help you?</div><div class="msg-time">\' + getTime() + \'</div></div></div>\';
        hideSuggestions();
        convId = null;
        ensureConversation();
    }

    // ========================================
    // SUGGESTIONS
    // ========================================
    function showSuggestions(list) {
        const row = document.getElementById(\'suggestionsRow\');
        if (!list || !list.length) { row.style.display = \'none\'; return; }
        row.innerHTML = list.map(s =>
            `<button class="suggestion-chip" onclick="useSuggestion(\'${s.replace(/\'/g, \\"\\'\\")}\')\">${escHtml(s)}</button>`
        ).join(\'\');
        row.style.display = \'flex\';
    }
    function hideSuggestions() {
        document.getElementById(\'suggestionsRow\').style.display = \'none\';
    }
    function useSuggestion(text) {
        document.getElementById(\'userInput\').value = text;
        hideSuggestions();
        sendMessage();
    }

    // ========================================
    // VOICE
    // ========================================
    function toggleVoice() {
        const btn = document.getElementById(\'micBtn\');
        const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
        if (!SR) { alert(\'Speech recognition not supported in this browser.\'); return; }
        if (isRecording) {
            recog && recog.stop();
            isRecording = false;
            btn.classList.remove(\'recording\');
            return;
        }
        recog = new SR();
        recog.lang = \'en-IN\';
        recog.interimResults = false;
        recog.onresult = e => { document.getElementById(\'userInput\').value = e.results[0][0].transcript; };
        recog.onend = () => { isRecording = false; btn.classList.remove(\'recording\'); };
        recog.onerror = () => { isRecording = false; btn.classList.remove(\'recording\'); };
        recog.start();
        isRecording = true;
        btn.classList.add(\'recording\');
    }

    // ========================================
    // ADD MESSAGE
    // ========================================
    function addMessage(html, type, opts) {
        opts = opts || {};
        const area = document.getElementById(\'messagesArea\');
        const msgEl = document.createElement(\'div\');
        msgEl.className = `message ${type}`;

        const wrap = document.createElement(\'div\');
        wrap.className = \'message-wrap\';

        if (opts.badge) {
            const b = document.createElement(\'div\');
            b.className = \'feature-badge\';
            b.textContent = opts.badge;
            wrap.appendChild(b);
        }
        const bubble = document.createElement(\'div\');
        bubble.className = \'message-bubble\';
        bubble.innerHTML = html;
        wrap.appendChild(bubble);

        const timeEl = document.createElement(\'div\');
        timeEl.className = \'msg-time\';
        timeEl.textContent = getTime();
        wrap.appendChild(timeEl);

        msgEl.appendChild(wrap);
        area.appendChild(msgEl);
        area.scrollTop = area.scrollHeight;
        return msgEl;
    }

    // ========================================
    // SEND MESSAGE
    // ========================================
    async function sendMessage() {
        const input = document.getElementById(\'userInput\');
        const message = input.value.trim();
        if (!message) return;
        input.value = \'\';
        hideSuggestions();
        await ensureConversation();

        addMessage(escHtml(message), \'user\');

        // build typing indicator
        const typEl = document.createElement(\'div\');
        typEl.className = \'message bot\';
        const guessed = selectedService || guessModule(message);
        const typLabel = guessed ? (SERVICE_NAMES[guessed] || guessed) : null;
        typEl.innerHTML = `<div class="message-wrap">
            <div class="typing-indicator">
                <div class="typing-dot"></div><div class="typing-dot"></div><div class="typing-dot"></div>
            </div>
            ${typLabel ? `<div class="typing-label">Querying ${typLabel}…</div>` : \'\'}
        </div>`;
        const area = document.getElementById(\'messagesArea\');
        area.appendChild(typEl);
        area.scrollTop = area.scrollHeight;

        try {
            const resp = await fetch(\'/ask\', {
                method: \'POST\',
                headers: { \'Content-Type\': \'application/json\' },
                body: JSON.stringify({ message, feature: selectedService, convId })
            });
            const data = await resp.json();
            typEl.remove();

            if (data.convId) convId = data.convId;
            const det = data.detected_feature;
            const badgeText = det ? `${MODULE_ICONS[det] || \'🤖\'} ${SERVICE_NAMES[det] || det}` : null;
            const isTabular = det === \'Timetable\' || det === \'Workload\';

            if (isTabular) {
                const parts = data.response.split(\'<br><br>\');
                let first = true;
                for (const part of parts) {
                    if (!part.trim()) continue;
                    addMessage(part, \'bot\', first && badgeText ? { badge: badgeText } : {});
                    first = false;
                    await new Promise(r => setTimeout(r, 380));
                }
            } else {
                addMessage(data.response, \'bot\', badgeText ? { badge: badgeText } : {});
            }

            if (data.suggestions && data.suggestions.length) showSuggestions(data.suggestions);

        } catch (err) {
            typEl.remove();
            addMessage(\'Sorry, an error occurred. Please try again.\', \'bot\');
        }
    }
</script>
</body>
</html>'''.replace('LOGO_PLACEHOLDER', LOGO)

with open('templates/index.html', 'w', encoding='utf-8') as f:
    f.write(HTML)

print("DONE - wrote", len(HTML), "chars")
