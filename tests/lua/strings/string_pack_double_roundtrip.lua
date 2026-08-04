-- vybe-test: lua/strings/string_pack_double_roundtrip
-- origin: languages/lua/tests/lua/test_strings.rs

local __w1 = "1.125\t9"
local __i = 0

local s = string.pack("d", 1.125)
do local __t = tostring(string.unpack("d", s)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
