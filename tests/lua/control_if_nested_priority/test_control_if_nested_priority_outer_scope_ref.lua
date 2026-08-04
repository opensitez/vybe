-- vybe-test: lua/control_if_nested_priority/test_control_if_nested_priority_outer_scope_ref
-- origin: languages/lua/tests/lua/test_control_if_nested_priority.rs

local __w1 = "1"
local __i = 0

local x=1; if true then local y=x; x='inner'; do local __t = tostring(y); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end; do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
