-- vybe-test: lua/math_library/math_fmod_keeps_fractional_remainder
-- origin: languages/lua/tests/lua/test_math_library.rs

local __w1 = "1.5"
local __i = 0

do local __t = tostring(math.fmod(5.5, 2)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
