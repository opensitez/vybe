-- vybe-test: lua/math_constants_rounding/math_pi_type
-- origin: languages/lua/tests/lua/test_math_constants_rounding.rs

local __w1 = "float"
local __i = 0

do local __t = tostring(math.type(math.pi)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
