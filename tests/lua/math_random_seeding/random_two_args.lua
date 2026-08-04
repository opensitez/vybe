-- vybe-test: lua/math_random_seeding/random_two_args
-- origin: languages/lua/tests/lua/test_math_random_seeding.rs

local __w1 = "true"
local __i = 0

math.randomseed(12345)
local r = math.random(5, 15)
do local __t = tostring(r >= 5 and r <= 15 and math.type(r) == "integer"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
