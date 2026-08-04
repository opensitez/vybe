-- vybe-test: lua/string_pack/pack_unpack_unsigned_byte
-- origin: languages/lua/tests/lua/test_string_pack.rs

local __w1 = "255\t2"
local __i = 0

local s=string.pack("B", 255)
do local __t = tostring(string.unpack("B", s)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
