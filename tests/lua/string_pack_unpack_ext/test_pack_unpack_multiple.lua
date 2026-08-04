-- vybe-test: lua/string_pack_unpack_ext/test_pack_unpack_multiple
-- origin: languages/lua/tests/lua/test_string_pack_unpack_ext.rs

local __w1 = "42 10 3 4"
local __i = 0

local s = string.pack('bBi', 42, 10, 3); local v1, v2, v3, p = string.unpack('bBi', s); do local __t = tostring(v1..' '..v2..' '..v3..' '..p); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
