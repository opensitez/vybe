-- vybe-test: lua/coroutine_status_running/test_running_with_arguments
-- origin: languages/lua/tests/lua/test_coroutine_status_running.rs

local __w1 = "true"
local __i = 0

local t = coroutine.create(function(x)
  local _, main = coroutine.running()
  return (main and 0 or x)
end)
local ok, v = coroutine.resume(t, 9)
do local __t = tostring(ok and v == 9); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
