-- 13_datetime_and_os_ops.lua
-- Demonstrates date/time formatting, timestamp math, and locale-aware formatting.

local now = os.time()
print("unix time:", now)
print("local date:", os.date("%Y-%m-%d %H:%M:%S", now))
print("utc date:", os.date("!%Y-%m-%d %H:%M:%S", now))

local event = {
  year = 2026,
  month = 12,
  day = 31,
  hour = 23,
  min = 0,
  sec = 0
}

local event_ts = os.time(event)
local diff = os.difftime(event_ts, now)
local days = math.floor(diff / (24 * 60 * 60))

print("days until event:", days)
print("weekday now:", os.date("%A"))
