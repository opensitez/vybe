-- vybe-test: lua/coroutines_extended/co_resume_yield_result
-- origin: languages/lua/tests/lua/test_coroutines_extended.rs

local __w1 = "99"
local __i = 0

local co = coroutine.create(function() local x = coroutine.yield() do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end)
coroutine.resume(co)  -- run to yield
coroutine.resume(co, 99)

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
