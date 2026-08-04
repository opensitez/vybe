-- vybe-test: lua/loops_for_negative_step/test_loops_for_negative_step_multiple_of_six
-- origin: languages/lua/tests/lua/test_loops_for_negative_step.rs

local __w1 = "36"
local __i = 0

local sum = 0
for i = 18, 0, -6 do sum = sum + i end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
