-- vybe-test: lua/string_pack_unpack_ext/test_pack_unpack_i8
-- origin: languages/lua/tests/lua/test_string_pack_unpack_ext.rs

local __w1 = "-1 9"
local __i = 0

local s = string.pack('i8', -1); local v, p = string.unpack('i8', s); do local __t = tostring(v..' '..p); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
