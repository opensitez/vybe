-- vybe-test: lua/math_random/test_math_random_invalid_args
-- origin: languages/lua/tests/lua/test_math_random.rs

local __w1 = "false"
local __i = 0

local ok = pcall(function() math.random(20, 10) end); do local __t = tostring(tostring(ok)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
