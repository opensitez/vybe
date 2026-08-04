-- vybe-test: lua/control_if_nested_priority/test_control_if_nested_priority_three_levels_outer
-- origin: languages/lua/tests/lua/test_control_if_nested_priority.rs

local __w1 = "neg"
local __i = 0

local x=-5; local y=''; if x > 0 then if x > 1 then if x > 10 then y='high' else y='mid' end else y='low' end else y='neg' end; do local __t = tostring(y); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
