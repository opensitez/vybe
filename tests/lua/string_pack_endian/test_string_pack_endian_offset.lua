-- vybe-test: lua/string_pack_endian/test_string_pack_endian_offset
-- origin: languages/lua/tests/lua/test_string_pack_endian.rs

local __w1 = "true"
local __i = 0

local s = string.pack("<i2", 9); do local __t = tostring(string.unpack("<i2", s) == 9); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
