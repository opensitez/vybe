-- vybe-test: lua/loops_while_break_control/test_loops_while_break_control_string_counter
-- origin: languages/lua/tests/lua/test_loops_while_break_control.rs

local __w1 = "123"
local __i = 0

local i = 0
local out = ''
while true do i = i + 1; out = out .. tostring(i); if i >= 3 then break end end
do local __t = tostring(out); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
