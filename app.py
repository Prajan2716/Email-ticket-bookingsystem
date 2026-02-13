"""
Flask Application - Runs sync on a schedule
"""
from flask import Flask, jsonify
from flask_apscheduler import APScheduler
from main import sync_mail_to_sheet
from datetime import datetime

app = Flask(__name__)

# Configure scheduler
class Config:
    SCHEDULER_API_ENABLED = True
    SCHEDULER_TIMEZONE = "UTC"

app.config.from_object(Config())

# Initialize scheduler
scheduler = APScheduler()
scheduler.init_app(app)

# Track sync status
sync_status = {
    "last_sync": None,
    "last_error": None,
    "sync_count": 0
}


@scheduler.task('interval', id='sync_job', seconds=5, misfire_grace_time=900)
def scheduled_sync():
    """Run sync every 5 seconds"""
    global sync_status
    try:
        print(f"\n⏰ Scheduled sync triggered at {datetime.now()}")
        sync_mail_to_sheet()
        sync_status["last_sync"] = datetime.now().isoformat()
        sync_status["last_error"] = None
        sync_status["sync_count"] += 1
    except Exception as e:
        sync_status["last_error"] = str(e)
        print(f"❌ Sync error: {e}")


@app.route('/')
def index():
    """Home page"""
    return """
    <html>
    <head>
        <title>Gmail Ticket System</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; }
            h1 { color: #333; }
            .status { background: #f0f0f0; padding: 20px; border-radius: 5px; }
            .running { color: green; font-weight: bold; }
        </style>
    </head>
    <body>
        <h1>📧 Gmail Ticket System</h1>
        <div class="status">
            <p class="running">✅ Scheduler is RUNNING</p>
            <p>Sync runs automatically every 5 seconds</p>
            <p><a href="/status">View Status</a></p>
        </div>
    </body>
    </html>
    """


@app.route('/status')
def status():
    """Get current sync status"""
    return jsonify({
        "scheduler_running": scheduler.running,
        "last_sync": sync_status["last_sync"],
        "last_error": sync_status["last_error"],
        "total_syncs": sync_status["sync_count"]
    })


if __name__ == '__main__':
    print("🚀 Starting Flask application with scheduler...")
    print("📱 Access the app at: http://localhost:5000")
    print("🔄 Sync will run automatically every 5 seconds")
    
    # Start scheduler
    scheduler.start()
    
    # Run Flask app
    app.run(debug=False, host='0.0.0.0', port=5000, use_reloader=False)