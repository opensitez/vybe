-- vybe-test: lua/loops_repeat_guard/test_loops_repeat_guard_with_function_call_condition
-- origin: languages/lua/tests/lua/test_loops_repeat_guard.rs

local __w1 = "2"
local __i = 0

local n = 0
local function done(v) return v >= 2 end
repeat n = n + 1 until done(n)
do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
