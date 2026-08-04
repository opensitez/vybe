-- vybe-test: lua/math_random_seed/test_math_random_seed_prefixed
-- origin: languages/lua/tests/lua/test_math_random_seed.rs

local __w1 = "true"
local __i = 0

math.randomseed(105)
local x = math.random()
math.randomseed(105)
local y = math.random()
do local __t = tostring(x == y); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
