-- vybe-test: lua/coroutine_status_running/test_running_nested_yield
-- origin: languages/lua/tests/lua/test_coroutine_status_running.rs

local __w1 = "true"
local __i = 0

local seen = false
local t = coroutine.create(function()
  local _, main = coroutine.running()
  coroutine.yield(main)
  seen = true
end)
coroutine.resume(t)
coroutine.resume(t)
do local __t = tostring(seen); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
