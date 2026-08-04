-- vybe-test: lua/coroutines_extended/co_resume_arguments
-- origin: languages/lua/tests/lua/test_coroutines_extended.rs

local __w1 = "xy"
local __i = 0

local co = coroutine.create(function(a, b) do local __t = tostring(a .. b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end)
coroutine.resume(co, "x", "y")

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
