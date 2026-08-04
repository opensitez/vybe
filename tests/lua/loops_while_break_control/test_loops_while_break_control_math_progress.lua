-- vybe-test: lua/loops_while_break_control/test_loops_while_break_control_math_progress
-- origin: languages/lua/tests/lua/test_loops_while_break_control.rs

local __w1 = "8"
local __i = 0

local x = 1
while x < 100 do x = x * 2; if x == 8 then break end end
do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
