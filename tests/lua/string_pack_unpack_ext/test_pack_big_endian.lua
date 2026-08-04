-- vybe-test: lua/string_pack_unpack_ext/test_pack_big_endian
-- origin: languages/lua/tests/lua/test_string_pack_unpack_ext.rs

local __w1 = "18 52"
local __i = 0

local s = string.pack('> H', 0x1234); local b1, b2 = string.unpack('B B', s); do local __t = tostring(b1..' '..b2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
