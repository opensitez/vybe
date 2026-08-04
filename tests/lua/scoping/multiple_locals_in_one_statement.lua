-- vybe-test: lua/scoping/multiple_locals_in_one_statement
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "7"
local __i = 0

local a, b = 3, 4
do local __t = tostring(a + b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
