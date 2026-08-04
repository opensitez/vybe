-- vybe-test: lua/loops_while_break_control/test_loops_while_break_control_last_value
-- origin: languages/lua/tests/lua/test_loops_while_break_control.rs

local __w1 = "9"
local __i = 0

local n = 0
while n < 6 do n = n + 1; if n == 5 then n = 9 end end
do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
