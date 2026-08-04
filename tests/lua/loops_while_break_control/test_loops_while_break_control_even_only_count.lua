-- vybe-test: lua/loops_while_break_control/test_loops_while_break_control_even_only_count
-- origin: languages/lua/tests/lua/test_loops_while_break_control.rs

local __w1 = "4"
local __i = 0

local i = 0
local even = 0
while i < 10 do i = i + 1; if i % 2 == 0 then even = even + 1 end if i == 9 then break end end
do local __t = tostring(even); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
