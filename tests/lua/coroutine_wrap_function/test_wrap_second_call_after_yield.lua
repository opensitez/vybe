-- vybe-test: lua/coroutine_wrap_function/test_wrap_second_call_after_yield
-- origin: languages/lua/tests/lua/test_coroutine_wrap_function.rs

local __w1 = "5"
local __i = 0

local f = coroutine.wrap(function() local x = coroutine.yield(1); return x + 1 end)
f()
do local __t = tostring(f(4)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
