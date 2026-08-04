-- vybe-test: lua/math_random_ext/test_math_random_same_range
-- origin: languages/lua/tests/lua/test_math_random_ext.rs

local __w1 = "5"
local __i = 0

local r = math.random(5, 5); do local __t = tostring(r); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
