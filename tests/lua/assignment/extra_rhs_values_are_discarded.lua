-- vybe-test: lua/assignment/extra_rhs_values_are_discarded
-- origin: languages/lua/tests/lua/test_assignment.rs

local __w1 = "1,2"
local __i = 0

local a, b = 1, 2, 3
do local __t = tostring(a .. "," .. b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
