-- vybe-test: lua/loops_for_zero_step_behavior/test_loops_for_zero_step_behavior_zero_step_never_entering_loop
-- origin: languages/lua/tests/lua/test_loops_for_zero_step_behavior.rs

local __w1 = "1"
local __i = 0

local total = 0
for i = 2, 2, -1 do total = total + 1 end
do local __t = tostring(total); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
