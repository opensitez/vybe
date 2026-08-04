-- vybe-test: lua/coroutines/coroutine_wrap_calls_function_directly
-- origin: languages/lua/tests/lua/test_coroutines.rs

local __w1 = "ok"
local __i = 0

local f = coroutine.wrap(function() return "ok" end)
do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
