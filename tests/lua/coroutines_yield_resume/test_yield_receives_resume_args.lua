-- vybe-test: lua/coroutines_yield_resume/test_yield_receives_resume_args
-- origin: languages/lua/tests/lua/test_coroutines_yield_resume.rs

local __w1 = "42"
local __i = 0

local co = coroutine.create(function() local x = coroutine.yield(); return x end); coroutine.resume(co); local ok, res = coroutine.resume(co, 42); do local __t = tostring(res); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
