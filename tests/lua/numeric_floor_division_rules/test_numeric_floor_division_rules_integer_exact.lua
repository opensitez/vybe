-- vybe-test: lua/numeric_floor_division_rules/test_numeric_floor_division_rules_integer_exact
-- origin: languages/lua/tests/lua/test_numeric_floor_division_rules.rs

local __w1 = "4"
local __i = 0

do local __t = tostring(12 // 3); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
