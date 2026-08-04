-- vybe-test: lua/loops_while_break_control/test_loops_while_break_control_break_and_continue_style
-- origin: languages/lua/tests/lua/test_loops_while_break_control.rs

local __w1 = "6"
local __i = 0

local i = 0
local sum = 0
while i < 6 do i = i + 1; if i == 4 then break end sum = sum + i end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
