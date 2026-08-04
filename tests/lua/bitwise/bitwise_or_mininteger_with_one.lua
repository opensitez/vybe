-- vybe-test: lua/bitwise/bitwise_or_mininteger_with_one
-- origin: languages/lua/tests/lua/test_bitwise.rs

local __w1 = "-9223372036854775807"
local __i = 0

do local __t = tostring(math.mininteger | 1); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
