-- vybe-test: lua/math_library/math_atan_of_zero
-- origin: languages/lua/tests/lua/test_math_library.rs

local __w1 = "0.0"
local __i = 0

do local __t = tostring(math.atan(0)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
