-- vybe-test: lua/math_random_ext/test_math_random_large
-- origin: languages/lua/tests/lua/test_math_random_ext.rs

local __w1 = "true"
local __i = 0

local r = math.random(1000000000); do local __t = tostring(type(r) == 'number' and r >= 1 and r <= 1000000000); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
