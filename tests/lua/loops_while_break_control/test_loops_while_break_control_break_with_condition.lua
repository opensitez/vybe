-- vybe-test: lua/loops_while_break_control/test_loops_while_break_control_break_with_condition
-- origin: languages/lua/tests/lua/test_loops_while_break_control.rs

local __w1 = "5"
local __i = 0

local count = 0
while count < 10 do count = count + 1; if count == 5 then break end end
do local __t = tostring(count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
