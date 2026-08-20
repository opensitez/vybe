-- App 39: Service Health Monitor

local services = {
  auth = {latency = {40, 45, 90, 120}, errors = {0, 0, 1, 2}},
  db = {latency = {12, 14, 16, 18}, errors = {0, 0, 0, 0}},
  api = {latency = {22, 23, 30, 28}, errors = {0, 1, 0, 0}},
}

local function avg(t)
  local s = 0
  for _, v in ipairs(t) do s = s + v end
  return s / #t
end

for name, m in pairs(services) do
  local lat = avg(m.latency)
  local err = avg(m.errors)
  local status = "healthy"
  if lat > 80 or err >= 1 then status = "degraded" end
  print(string.format("%s latency=%.1f errors=%.2f status=%s", name, lat, err, status))
end
