-- vybe-test: lua/loops_for_zero_step_behavior/test_loops_for_zero_step_behavior_step_expression_zero_guarded
-- origin: languages/lua/tests/lua/test_loops_for_zero_step_behavior.rs

local __w1 = "3"
local __i = 0

local count = 0
local step = 0
local run_zero = false
if run_zero then for i = 3, 1, step do count = count + 1 end else for i = 3, 1, -1 do count = count + 1 end end
do local __t = tostring(count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
