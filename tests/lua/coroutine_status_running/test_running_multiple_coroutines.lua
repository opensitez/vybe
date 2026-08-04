-- vybe-test: lua/coroutine_status_running/test_running_multiple_coroutines
-- origin: languages/lua/tests/lua/test_coroutine_status_running.rs

local __w1 = "true"
local __i = 0

local c1 = coroutine.create(function()
  local th = coroutine.running()
  return th ~= nil
end)
local c2 = coroutine.create(function()
  local th = coroutine.running()
  return th ~= nil
end)
local _, a = coroutine.resume(c1)
local _, b = coroutine.resume(c2)
do local __t = tostring(a == true and b == true); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
