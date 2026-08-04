-- vybe-test: lua/coroutines/coroutine_wrap_propagates_yield_values
-- origin: languages/lua/tests/lua/test_coroutines.rs

local __w1 = "3"
local __i = 0

local f = coroutine.wrap(function() coroutine.yield(3) return 9 end)
do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
