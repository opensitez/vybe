-- vybe-test: lua/loops_for_negative_step/test_loops_for_negative_step_with_local_variable_step
-- origin: languages/lua/tests/lua/test_loops_for_negative_step.rs

local __w1 = "40"
local __i = 0

local sum = 0
local step = -4
for i = 16, 1, step do sum = sum + i end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
