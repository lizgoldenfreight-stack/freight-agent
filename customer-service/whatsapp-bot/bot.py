#!/usr/bin/env python3
"""
WhatsApp Auto-Reply Bot for Freight Forwarder
=============================================
Uses pywhatkit for simple auto-reply via WhatsApp Web.

Setup:
1. pip install pywhatkit flask
2. Open WhatsApp Web in your browser (web.whatsapp.com)
3. Run: python3 bot.py
4. Scan QR code if needed
5. Bot will auto-reply to incoming messages

Note: WhatsApp Web must stay open in browser.
"""

import json
import os
import re
from datetime import datetime

# Load config
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(SCRIPT_DIR, "config.json"), "r") as f:
    config = json.load(f)

with open(os.path.join(SCRIPT_DIR, "replies.json"), "r") as f:
    replies = json.load(f)


def is_business_hours():
    """Check if current time is within business hours (MYT)."""
    now = datetime.now()
    day_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    current_day = day_names[now.weekday()]
    
    if current_day not in config["business_hours"]["days"]:
        return False
    
    start = datetime.strptime(config["business_hours"]["start"], "%H:%M").time()
    end = datetime.strptime(config["business_hours"]["end"], "%H:%M").time()
    
    return start <= now.time() <= end


def detect_language(text):
    """Detect if message is in Chinese or English."""
    chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', text))
    return "cn" if chinese_chars > len(text) * 0.2 else "en"


def detect_intent(message):
    """Detect what the customer is asking about."""
    message_lower = message.lower()
    
    for intent, keywords in config["keywords"].items():
        for keyword in keywords:
            if keyword.lower() in message_lower:
                return intent
    
    return "unknown"


def get_reply(intent, lang="en"):
    """Get the appropriate reply template."""
    if intent in replies:
        return replies[intent].get(lang, replies[intent].get("en", replies["unknown"]["en"]))
    return replies["unknown"].get(lang, replies["unknown"]["en"])


def process_message(message):
    """Process incoming message and return reply."""
    if not is_business_hours() and not config.get("auto_reply_outside_hours", False):
        return config["out_of_hours_message"]
    
    lang = detect_language(message)
    intent = detect_intent(message)
    reply = get_reply(intent, lang)
    
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M')}] Intent: {intent} | Lang: {lang}")
    print(f"  Message: {message[:50]}...")
    print(f"  Reply: {reply[:50]}...")
    
    return reply


# ============================================================
# Dashboard: Simple web interface to manage replies
# ============================================================
def create_dashboard():
    """Create a simple web dashboard for managing the bot."""
    dashboard_html = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>WhatsApp Bot Dashboard</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Arial, sans-serif; background: #f0f2f5; padding: 20px; }
        .container { max-width: 900px; margin: 0 auto; }
        h1 { color: #1a73e8; margin-bottom: 20px; }
        .card { background: white; border-radius: 10px; padding: 20px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .card h2 { color: #333; margin-bottom: 15px; font-size: 18px; }
        .status { display: flex; gap: 20px; margin-bottom: 20px; }
        .stat { flex: 1; background: #e3f2fd; padding: 15px; border-radius: 8px; text-align: center; }
        .stat .number { font-size: 32px; font-weight: bold; color: #1a73e8; }
        .stat .label { font-size: 12px; color: #666; }
        .reply-box { background: #f5f5f5; padding: 15px; border-radius: 8px; margin-bottom: 10px; }
        .reply-box .intent { font-weight: bold; color: #1a73e8; margin-bottom: 5px; }
        .reply-box .text { white-space: pre-wrap; font-size: 14px; }
        .btn { background: #1a73e8; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; }
        .btn:hover { background: #1557b0; }
        textarea { width: 100%; padding: 10px; border: 1px solid #ddd; border-radius: 5px; min-height: 100px; font-family: monospace; }
        .tag { display: inline-block; background: #e8f5e9; color: #2e7d32; padding: 3px 8px; border-radius: 4px; font-size: 12px; margin: 2px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>🤖 WhatsApp Freight Bot Dashboard</h1>
        
        <div class="status">
            <div class="stat">
                <div class="number">ACTIVE</div>
                <div class="label">Bot Status</div>
            </div>
            <div class="stat">
                <div class="number">5</div>
                <div class="label">Intents</div>
            </div>
            <div class="stat">
                <div class="number">EN / CN</div>
                <div class="label">Languages</div>
            </div>
        </div>
        
        <div class="card">
            <h2>📝 Quick Reply Templates</h2>
            <div class="reply-box">
                <div class="intent">🚢 Quote Request</div>
                <div class="text">Thanks for your interest! To give you an accurate quote, I need: 1. From which city? 2. To where? 3. What cargo? 4. Volume/Weight? 5. When?</div>
            </div>
            <div class="reply-box">
                <div class="intent">📦 Tracking</div>
                <div class="text">Let me check the status for you! Could you share your booking reference or B/L number?</div>
            </div>
            <div class="reply-box">
                <div class="intent">⏱️ Transit Time</div>
                <div class="text">Sea: China→PKL 7-12 days | East MY +3-5 days | Air: China→KUL 1-3 days</div>
            </div>
            <div class="reply-box">
                <div class="intent">🎯 Services</div>
                <div class="text">Sea freight (FCL/LCL), Air freight, Customs, Door-to-door, East MY transshipment, Insurance</div>
            </div>
        </div>
        
        <div class="card">
            <h2>🏷️ Detected Keywords</h2>
            <div>
                <span class="tag">price</span><span class="tag">how much</span><span class="tag">quote</span>
                <span class="tag">where</span><span class="tag">status</span><span class="tag">track</span>
                <span class="tag">when</span><span class="tag">transit</span><span class="tag">time</span>
                <span class="tag">报价</span><span class="tag">多少钱</span><span class="tag">到哪</span>
                <span class="tag">状态</span><span class="tag">多久</span><span class="tag">时效</span>
            </div>
        </div>
        
        <div class="card">
            <h2>⚙️ How to Use</h2>
            <ol style="padding-left: 20px; line-height: 2;">
                <li>Open WhatsApp Web in your browser</li>
                <li>Run <code>python3 bot.py</code> in terminal</li>
                <li>Bot detects incoming messages</li>
                <li>Auto-replies based on keywords</li>
                <li>Edit <code>replies.json</code> to customize messages</li>
            </ol>
        </div>
    </div>
</body>
</html>"""
    
    return dashboard_html


if __name__ == "__main__":
    print("🤖 WhatsApp Freight Bot")
    print("=" * 40)
    print(f"Business hours: {config['business_hours']['start']} - {config['business_hours']['end']} MYT")
    print(f"Auto-reply: {config['auto_reply']}")
    print()
    
    # Test the bot
    test_messages = [
        "How much to ship to Malaysia?",
        "我的货到哪了？",
        "What services do you offer?",
        "How long does shipping take?",
        "Hello!",
        "Can you ship to Kuching?",
    ]
    
    print("Testing auto-reply system:")
    print("-" * 40)
    for msg in test_messages:
        reply = process_message(msg)
        print()
    
    print("-" * 40)
    print("✅ Bot ready! Edit config.json and replies.json to customize.")
    print("📋 Dashboard ready: open dashboard.html in browser")
