-- vybe-test: lua/math_library_extended/math_max_mixed
-- origin: languages/lua/tests/lua/test_math_library_extended.rs

local __w1 = "5"
local __i = 0

do local __t = tostring(math.max(-5, 0, 5)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
