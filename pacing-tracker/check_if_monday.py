from datetime import datetime
from zoneinfo import ZoneInfo

now = datetime.now(ZoneInfo("America/Chicago"))
if now.weekday() != 0:
    raise SystemExit(f"Not Monday — current CST day is {now.strftime('%A')}.")
