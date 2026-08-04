-- vybe-test: lua/math_abs_cases/test_math_abs_cases_edge_second
-- origin: languages/lua/tests/lua/test_math_abs_cases.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(math.abs(17) == 17); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
