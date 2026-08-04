-- vybe-test: lua/string_pack_unpack_ext/test_pack_unpack_f
-- origin: languages/lua/tests/lua/test_string_pack_unpack_ext.rs

local __w1 = "true"
local __i = 0

local s = string.pack('f', 3.14); local v, p = string.unpack('f', s); do local __t = tostring(tostring(math.abs(v - 3.14) < 0.01)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
