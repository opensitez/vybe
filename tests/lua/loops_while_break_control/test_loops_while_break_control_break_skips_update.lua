-- vybe-test: lua/loops_while_break_control/test_loops_while_break_control_break_skips_update
-- origin: languages/lua/tests/lua/test_loops_while_break_control.rs

local __w1 = "3:20"
local __i = 0

local value = 0
local i = 0
while i < 5 do i = i + 1; if i == 3 then break end value = value + 10 end
do local __t = tostring(i .. ':' .. value); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
