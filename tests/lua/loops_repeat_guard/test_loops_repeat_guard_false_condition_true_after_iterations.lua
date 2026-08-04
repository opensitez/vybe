-- vybe-test: lua/loops_repeat_guard/test_loops_repeat_guard_false_condition_true_after_iterations
-- origin: languages/lua/tests/lua/test_loops_repeat_guard.rs

local __w1 = "6"
local __i = 0

local sum = 0
repeat sum = sum + 2; local next = sum > 4; until next
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
