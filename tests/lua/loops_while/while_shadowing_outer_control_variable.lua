-- vybe-test: lua/loops_while/while_shadowing_outer_control_variable
-- origin: languages/lua/tests/lua/test_loops_while.rs

local __w1 = "5"
local __i = 0

local i = 10
while i > 0 do
  local i = 5
  do local __t = tostring(i); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
  break
end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
