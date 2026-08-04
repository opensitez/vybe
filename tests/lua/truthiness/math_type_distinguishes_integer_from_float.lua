-- vybe-test: lua/truthiness/math_type_distinguishes_integer_from_float
-- origin: languages/lua/tests/lua/test_truthiness.rs

local __w1 = "integer,float"
local __i = 0

do local __t = tostring(math.type(1) .. ',' .. math.type(1.0)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
