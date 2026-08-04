-- vybe-test: lua/loops_for_negative_step/test_loops_for_negative_step_decrement_with_multiplier
-- origin: languages/lua/tests/lua/test_loops_for_negative_step.rs

local __w1 = "384"
local __i = 0

local out = 1
for i = 8, 2, -2 do out = out * i end
do local __t = tostring(out); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
