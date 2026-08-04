-- vybe-test: lua/loops_for_zero_step_behavior/test_loops_for_zero_step_behavior_conditional_zero_step
-- origin: languages/lua/tests/lua/test_loops_for_zero_step_behavior.rs

local __w1 = "empty"
local __i = 0

local count = 0
if true then for i = 4, 2, 0 do count = count + 1 end end
do local __t = tostring(count == 0 and "empty" or "nonempty"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
