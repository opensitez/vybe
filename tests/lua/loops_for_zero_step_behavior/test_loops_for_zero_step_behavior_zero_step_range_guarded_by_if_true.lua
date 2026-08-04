-- vybe-test: lua/loops_for_zero_step_behavior/test_loops_for_zero_step_behavior_zero_step_range_guarded_by_if_true
-- origin: languages/lua/tests/lua/test_loops_for_zero_step_behavior.rs

local __w1 = "empty"
local __i = 0

local value = ""
if 1 > 2 then for i = 1, 0, 0 do value = value .. "x" end end
do local __t = tostring(value == "" and "empty" or "filled"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
