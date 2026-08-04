-- vybe-test: lua/control_flow/while_with_break_on_condition
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "3"
local __i = 0

local i = 0
while true do
  i = i + 1
  if i >= 3 then break end
end
do local __t = tostring(i); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
