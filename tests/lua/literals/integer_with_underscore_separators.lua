-- vybe-test: lua/literals/integer_with_underscore_separators
-- origin: languages/lua/tests/lua/test_literals.rs

local __w1 = "3000"
local __i = 0

do local __t = tostring(1_000 + 2_000); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
