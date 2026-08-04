-- vybe-test: lua/loops_for_zero_step_behavior/test_loops_for_zero_step_behavior_local_bounds_areolated
-- origin: languages/lua/tests/lua/test_loops_for_zero_step_behavior.rs

local __w1 = "0"
local __i = 0

local start = 9
local stop = 3
local sum = 0
if start > stop then for i = start, stop, 0 do sum = sum + 1 end end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
