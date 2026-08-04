-- vybe-test: lua/truthiness/and_returns_last_value_when_all_truthy
-- origin: languages/lua/tests/lua/test_truthiness.rs

local __w1 = "3"
local __i = 0

do local __t = tostring(1 and 2 and 3); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
