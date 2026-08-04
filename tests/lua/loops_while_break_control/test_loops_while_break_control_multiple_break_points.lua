-- vybe-test: lua/loops_while_break_control/test_loops_while_break_control_multiple_break_points
-- origin: languages/lua/tests/lua/test_loops_while_break_control.rs

local __w1 = "21"
local __i = 0

local i = 0
local total = 0
while i < 20 do i = i + 2; if i == 4 then total = total + 1 elseif i == 10 then break end total = total + i end
do local __t = tostring(total); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
