-- vybe-test: lua/coroutines_basics/test_coroutine_resume_args
-- origin: languages/lua/tests/lua/test_coroutines_basics.rs

local __w1 = "true 30"
local __i = 0

local co = coroutine.create(function(a,b) return a+b end); local ok, res = coroutine.resume(co, 10, 20); do local __t = tostring(tostring(ok)..' '..res); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
