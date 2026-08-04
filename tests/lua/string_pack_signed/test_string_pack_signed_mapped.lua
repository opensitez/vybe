-- vybe-test: lua/string_pack_signed/test_string_pack_signed_mapped
-- origin: languages/lua/tests/lua/test_string_pack_signed.rs

local __w1 = "true"
local __i = 0

local s = string.pack("i4", 14); local v = string.unpack("i4", s); do local __t = tostring(v == 14); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
