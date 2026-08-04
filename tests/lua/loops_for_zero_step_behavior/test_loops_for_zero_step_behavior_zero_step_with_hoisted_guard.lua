-- vybe-test: lua/loops_for_zero_step_behavior/test_loops_for_zero_step_behavior_zero_step_with_hoisted_guard
-- origin: languages/lua/tests/lua/test_loops_for_zero_step_behavior.rs

local __w1 = "0"
local __i = 0

local value = 0
local start = 0
local stop = 1
if stop < start then for i = start, stop, 0 do value = value + 1 end end
do local __t = tostring(value); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
