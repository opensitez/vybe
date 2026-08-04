-- vybe-test: lua/math_random_range/test_math_random_range_nested
-- origin: languages/lua/tests/lua/test_math_random_range.rs

local __w1 = "true"
local __i = 0

math.randomseed(15)
local x = math.random(1, 13)
do local __t = tostring(x >= 1 and x <= 13); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
