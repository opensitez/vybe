-- vybe-test: lua/math_random/test_math_random_one_arg
-- origin: languages/lua/tests/lua/test_math_random.rs

local __w1 = "true"
local __i = 0

local r = math.random(5); do local __t = tostring(type(r) == 'number' and r >= 1 and r <= 5 and math.floor(r) == r); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
