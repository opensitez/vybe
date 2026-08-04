-- vybe-test: lua/globals/undeclared_assignment_creates_global
-- origin: languages/lua/tests/lua/test_globals.rs

local __w1 = "10"
local __i = 0

foo = 10
do local __t = tostring(foo); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
