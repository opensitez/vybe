-- vybe-test: lua/base_tonumber_radix/test_tonumber_radix_edge_second
-- origin: languages/lua/tests/lua/test_base_tonumber_radix.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(tonumber("2a", 11) == 32); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
