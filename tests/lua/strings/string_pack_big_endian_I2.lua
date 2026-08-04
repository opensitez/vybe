-- vybe-test: lua/strings/string_pack_big_endian_I2
-- origin: languages/lua/tests/lua/test_strings.rs

local __w1 = "1000\t3"
local __i = 0

local s = string.pack(">I2", 1000)
do local __t = tostring(string.unpack(">I2", s)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
