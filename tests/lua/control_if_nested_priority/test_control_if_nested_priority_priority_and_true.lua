-- vybe-test: lua/control_if_nested_priority/test_control_if_nested_priority_priority_and_true
-- origin: languages/lua/tests/lua/test_control_if_nested_priority.rs

local __w1 = "ok"
local __i = 0

local x=2; local y=''; if x > 0 and x < 10 and x == 2 then y='ok' else y='bad' end; do local __t = tostring(y); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
