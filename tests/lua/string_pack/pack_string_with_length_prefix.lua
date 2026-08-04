-- vybe-test: lua/string_pack/pack_string_with_length_prefix
-- origin: languages/lua/tests/lua/test_string_pack.rs

local __w1 = "hi\t4"
local __i = 0

local s=string.pack("z", "hi")
do local __t = tostring(string.unpack("z", s)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
