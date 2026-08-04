-- vybe-test: lua/loops_while_break_control/test_loops_while_break_control_break_after_nested_operation
-- origin: languages/lua/tests/lua/test_loops_while_break_control.rs

local __w1 = "7"
local __i = 0

local i = 0
local total = 0
while true do i = i + 1; if i == 1 then total = total + 1 else total = total + 2 end if total > 5 then break end end
do local __t = tostring(total); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
