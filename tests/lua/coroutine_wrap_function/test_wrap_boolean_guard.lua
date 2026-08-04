-- vybe-test: lua/coroutine_wrap_function/test_wrap_boolean_guard
-- origin: languages/lua/tests/lua/test_coroutine_wrap_function.rs

local __w1 = "true"
local __i = 0

local f = coroutine.wrap(function() return not false end)
do local __t = tostring(f() == true); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
