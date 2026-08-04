-- vybe-test: lua/coroutines_yield_resume/test_yield_resume_multiple_values
-- origin: languages/lua/tests/lua/test_coroutines_yield_resume.rs

local __w1 = "30"
local __i = 0

local co = coroutine.create(function() local a, b = coroutine.yield(); return a+b end); coroutine.resume(co); local ok, res = coroutine.resume(co, 10, 20); do local __t = tostring(res); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
