-- vybe-test: lua/string_pack/pack_fixed_string_with_c
-- origin: languages/lua/tests/lua/test_string_pack.rs

local __w1 = "lua!\t6"
local __i = 0

local s=string.pack("z", "lua!")
do local __t = tostring(string.unpack("z", s)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
