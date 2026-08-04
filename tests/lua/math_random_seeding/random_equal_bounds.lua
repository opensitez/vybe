-- vybe-test: lua/math_random_seeding/random_equal_bounds
-- origin: languages/lua/tests/lua/test_math_random_seeding.rs

local __w1 = "7"
local __i = 0

do local __t = tostring(math.random(7, 7)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
