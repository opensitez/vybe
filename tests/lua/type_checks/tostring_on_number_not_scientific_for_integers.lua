-- vybe-test: lua/type_checks/tostring_on_number_not_scientific_for_integers
-- origin: languages/lua/tests/lua/test_type_checks.rs

local __w1 = "1000"
local __i = 0

do local __t = tostring(tostring(1000)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
