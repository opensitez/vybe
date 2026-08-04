-- vybe-test: lua/control_flow/while_with_or_condition_evaluates_second_part
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "true"
local __i = 0

local a, b = false, true
local ran = false
while a or b do
  ran = true
  b = false
end
do local __t = tostring(ran); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
