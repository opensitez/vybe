-- vybe-test: lua/operators/string_number_less_than_does_not_coerce
-- origin: languages/lua/tests/lua/test_operators.rs

local __w1 = "false"
local __i = 0

do local __t = tostring("2" < 12); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
