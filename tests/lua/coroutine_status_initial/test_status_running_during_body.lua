-- vybe-test: lua/coroutine_status_initial/test_status_running_during_body
-- origin: languages/lua/tests/lua/test_coroutine_status_initial.rs

local __w1 = "true"
local __i = 0

local active = false
local t = coroutine.create(function()
  active = (coroutine.running() ~= nil)
end)
coroutine.resume(t)
do local __t = tostring(active); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
