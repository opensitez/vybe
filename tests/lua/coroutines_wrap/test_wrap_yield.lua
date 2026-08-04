-- vybe-test: lua/coroutines_wrap/test_wrap_yield
-- origin: languages/lua/tests/lua/test_coroutines_wrap.rs

local __w1 = "1 2"
local __i = 0

local f = coroutine.wrap(function() coroutine.yield(1) return 2 end); do local __t = tostring(f()..' '..f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
