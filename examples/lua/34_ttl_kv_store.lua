-- App 34: In-Memory KV Store with TTL

local store = {}
local now = 0

local function set(key, value, ttl)
  store[key] = {value = value, exp = ttl and (now + ttl) or math.huge}
end

local function get(key)
  local e = store[key]
  if not e then return nil end
  if now >= e.exp then
    store[key] = nil
    return nil
  end
  return e.value
end

local function tick(dt)
  now = now + dt
end

set("session:1", "alice", 3)
set("cfg:theme", "ocean")

for t = 0, 5 do
  print("t=" .. now, "session", get("session:1"), "theme", get("cfg:theme"))
  tick(1)
end
