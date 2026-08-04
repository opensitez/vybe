-- vybe-test: lua/math_floor_cases/test_math_floor_cases_negative
-- origin: languages/lua/tests/lua/test_math_floor_cases.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(math.floor(6.2 + 6) == math.floor(12.2)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
