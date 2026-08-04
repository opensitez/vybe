-- vybe-test: lua/literals/true_and_false_in_expression
-- origin: languages/lua/tests/lua/test_literals.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(true and false or true); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
