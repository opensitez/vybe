-- vybe-test: lua/loops_while_break_control/test_loops_while_break_control_break_immediately
-- origin: languages/lua/tests/lua/test_loops_while_break_control.rs

local __w1 = "0"
local __i = 0

local sum = 0
local i = 0
while true do break end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
