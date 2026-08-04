-- vybe-test: lua/coroutines_extended/co_wrap_callable
-- origin: languages/lua/tests/lua/test_coroutines_extended.rs

local __w1 = "11"
local __i = 0

local f = coroutine.wrap(function(x) return x + 1 end)
do local __t = tostring(f(10)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
