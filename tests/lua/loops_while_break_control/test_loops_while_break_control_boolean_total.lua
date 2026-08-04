-- vybe-test: lua/loops_while_break_control/test_loops_while_break_control_boolean_total
-- origin: languages/lua/tests/lua/test_loops_while_break_control.rs

local __w1 = "true"
local __i = 0

local i = 0
local ok = false
while i < 4 do i = i + 1; if i == 2 then ok = true end if i == 3 then break end end
do local __t = tostring(ok and "true" or "false"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
