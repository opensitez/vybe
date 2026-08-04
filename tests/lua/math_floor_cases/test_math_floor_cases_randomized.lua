-- vybe-test: lua/math_floor_cases/test_math_floor_cases_randomized
-- origin: languages/lua/tests/lua/test_math_floor_cases.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(math.floor(18.2 + 18) == math.floor(36.2)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
