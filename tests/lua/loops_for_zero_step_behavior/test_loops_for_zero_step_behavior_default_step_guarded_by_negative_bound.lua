-- vybe-test: lua/loops_for_zero_step_behavior/test_loops_for_zero_step_behavior_default_step_guarded_by_negative_bound
-- origin: languages/lua/tests/lua/test_loops_for_zero_step_behavior.rs

local __w1 = "0"
local __i = 0

local count = 0
for i = 5, 1 do count = count + 1 end
do local __t = tostring(count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
