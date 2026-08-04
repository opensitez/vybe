-- vybe-test: lua/control_if_nested_priority/test_control_if_nested_priority_inner_else
-- origin: languages/lua/tests/lua/test_control_if_nested_priority.rs

local __w1 = "inner_else"
local __i = 0

local x=1; local y=''; if x > 0 then if x < 0 then y='inner' else y='inner_else' end else y='outer' end; do local __t = tostring(y); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
