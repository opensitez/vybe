-- vybe-test: lua/loops_while_break_control/test_loops_while_break_control_while_false_after_break_guard
-- origin: languages/lua/tests/lua/test_loops_while_break_control.rs

local __w1 = "1"
local __i = 0

local active = true
local count = 0
while active do count = count + 1; if count > 2 then active = false end if count == 1 then break end end
do local __t = tostring(count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
