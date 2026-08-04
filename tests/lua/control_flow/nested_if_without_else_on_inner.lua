-- vybe-test: lua/control_flow/nested_if_without_else_on_inner
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "2"
local __i = 0

if true then if false then do local __t = tostring(1); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end do local __t = tostring(2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
