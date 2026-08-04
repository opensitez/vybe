-- vybe-test: lua/numeric_floor_division_rules/test_numeric_floor_division_rules_nested
-- origin: languages/lua/tests/lua/test_numeric_floor_division_rules.rs

local __w1 = "2"
local __i = 0

do local __t = tostring((20 // 5) // 2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
