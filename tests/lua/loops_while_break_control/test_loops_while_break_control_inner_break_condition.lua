-- vybe-test: lua/loops_while_break_control/test_loops_while_break_control_inner_break_condition
-- origin: languages/lua/tests/lua/test_loops_while_break_control.rs

local __w1 = "8"
local __i = 0

local sum = 0
local i = 0
while i < 100 do i = i + 1; if i % 3 == 0 then if i > 6 then break end end sum = sum + 1 end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
