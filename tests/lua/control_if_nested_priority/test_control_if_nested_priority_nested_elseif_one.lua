-- vybe-test: lua/control_if_nested_priority/test_control_if_nested_priority_nested_elseif_one
-- origin: languages/lua/tests/lua/test_control_if_nested_priority.rs

local __w1 = "one"
local __i = 0

local x=1; local y=''; if x < 0 then y='neg' elseif x == 0 then y='zero' elseif x == 1 then y='one' else y='other' end; do local __t = tostring(y); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
