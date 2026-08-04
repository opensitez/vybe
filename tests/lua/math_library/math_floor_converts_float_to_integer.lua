-- vybe-test: lua/math_library/math_floor_converts_float_to_integer
-- origin: languages/lua/tests/lua/test_math_library.rs

local __w1 = "9"
local __i = 0

do local __t = tostring(math.floor(9.9)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
