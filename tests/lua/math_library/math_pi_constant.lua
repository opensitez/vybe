-- vybe-test: lua/math_library/math_pi_constant
-- origin: languages/lua/tests/lua/test_math_library.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(math.pi > 3 and math.pi < 4); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
