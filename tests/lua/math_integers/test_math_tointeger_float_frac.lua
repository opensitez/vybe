-- vybe-test: lua/math_integers/test_math_tointeger_float_frac
-- origin: languages/lua/tests/lua/test_math_integers.rs

local __w1 = "nil"
local __i = 0

do local __t = tostring(math.tointeger(10.5) or 'nil'); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
