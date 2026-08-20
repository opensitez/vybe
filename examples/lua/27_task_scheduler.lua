-- App 27: Cooperative Task Scheduler

local Scheduler = {}
Scheduler.__index = Scheduler

function Scheduler.new()
  return setmetatable({queue = {}, now = 0}, Scheduler)
end

function Scheduler:add(delay, fn)
  self.queue[#self.queue + 1] = {at = self.now + delay, fn = fn}
end

function Scheduler:run(until_time)
  table.sort(self.queue, function(a, b) return a.at < b.at end)
  while #self.queue > 0 do
    local task = table.remove(self.queue, 1)
    if task.at > until_time then break end
    self.now = task.at
    task.fn(self)
    table.sort(self.queue, function(a, b) return a.at < b.at end)
  end
end

local s = Scheduler.new()
local retries = 0

s:add(1, function(sc)
  print("t=" .. sc.now, "poll queue")
  sc:add(2, function(sc2) print("t=" .. sc2.now, "poll queue again") end)
end)

s:add(2, function(sc)
  retries = retries + 1
  print("t=" .. sc.now, "send email attempt", retries)
  if retries < 3 then sc:add(2, debug.getinfo(1, "f").func) end
end)

s:run(10)
