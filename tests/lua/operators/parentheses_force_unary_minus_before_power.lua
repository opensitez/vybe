-- vybe-test: lua/operators/parentheses_force_unary_minus_before_power
-- origin: languages/lua/tests/lua/test_operators.rs

local __w1 = "4.0"
local __i = 0

do local __t = tostring((-2) ^ 2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
