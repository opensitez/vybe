-- vybe-test: lua/base_tonumber_radix/test_tonumber_radix_baseline
-- origin: languages/lua/tests/lua/test_base_tonumber_radix.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(tonumber("1111", 2) == 15); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
