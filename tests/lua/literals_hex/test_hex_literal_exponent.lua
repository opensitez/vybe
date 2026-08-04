-- vybe-test: lua/literals_hex/test_hex_literal_exponent
-- origin: languages/lua/tests/lua/test_literals_hex.rs

local __w1 = "16.0"
local __i = 0

do local __t = tostring(0x1p4); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
