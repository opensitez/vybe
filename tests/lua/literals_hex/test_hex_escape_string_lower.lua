-- vybe-test: lua/literals_hex/test_hex_escape_string_lower
-- origin: languages/lua/tests/lua/test_literals_hex.rs

local __w1 = "abc"
local __i = 0

do local __t = tostring('\x61\x62\x63'); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
