-- vybe-test: lua/loops_while_break_control/test_loops_while_break_control_flag_toggle
-- origin: languages/lua/tests/lua/test_loops_while_break_control.rs

local __w1 = "2"
local __i = 0

local run = true
local i = 0
while run do i = i + 1; if i == 1 then i = i + 1 end; if i > 3 then break end; run = false end
do local __t = tostring(i); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
