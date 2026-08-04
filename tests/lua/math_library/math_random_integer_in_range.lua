-- vybe-test: lua/math_library/math_random_integer_in_range
-- origin: languages/lua/tests/lua/test_math_library.rs

local __w1 = "true"
local __i = 0

math.randomseed(2)
local r = math.random(3, 5)
do local __t = tostring(r >= 3 and r <= 5); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
