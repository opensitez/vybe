-- vybe-test: lua/control_flow/while_repeats_until_condition_false
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "2"
local __i = 0

local i = 0
while i < 2 do i = i + 1 end
do local __t = tostring(i); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
