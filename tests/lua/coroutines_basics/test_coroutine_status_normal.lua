-- vybe-test: lua/coroutines_basics/test_coroutine_status_normal
-- origin: languages/lua/tests/lua/test_coroutines_basics.rs

local __w1 = "running"
local __i = 0

local st; local co1 = coroutine.create(function() st=coroutine.status(coroutine.running()) end); local co2 = coroutine.create(function() coroutine.resume(co1) end); coroutine.resume(co2); do local __t = tostring(st); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
