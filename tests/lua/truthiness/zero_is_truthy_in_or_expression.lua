-- vybe-test: lua/truthiness/zero_is_truthy_in_or_expression
-- origin: languages/lua/tests/lua/test_truthiness.rs

local __w1 = "true"
local __i = 0

do local __t = tostring((0 or 99) == 0); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
