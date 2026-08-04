-- vybe-test: lua/base_assert_variadics/test_assert_varargs_decimal
-- origin: languages/lua/tests/lua/test_base_assert_variadics.rs

local __w1 = "true"
local __i = 0

local v1, v2, v3 = assert(4, 5, 6); do local __t = tostring(v1 + v2 == 4 + 5); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
