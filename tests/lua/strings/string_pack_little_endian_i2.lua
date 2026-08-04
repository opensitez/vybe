-- vybe-test: lua/strings/string_pack_little_endian_i2
-- origin: languages/lua/tests/lua/test_strings.rs

local __w1 = "-1000\t3"
local __i = 0

local s = string.pack("<i2", -1000)
do local __t = tostring(string.unpack("<i2", s)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
