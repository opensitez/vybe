-- vybe-test: lua/loops_while_break_control/test_loops_while_break_control_no_break_to_end
-- origin: languages/lua/tests/lua/test_loops_while_break_control.rs

local __w1 = "10"
local __i = 0

local sum = 0
local i = 1
while i <= 4 do sum = sum + i; i = i + 1 end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
