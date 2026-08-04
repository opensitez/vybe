-- vybe-test: lua/scoping/scoping_repeat_until_condition_can_access_locals_in_loop_body
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "10"
local __i = 0

local count = 0
repeat
  local x = 10
  count = count + x
until x == 10
do local __t = tostring(count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
