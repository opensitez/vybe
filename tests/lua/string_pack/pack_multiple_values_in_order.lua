-- vybe-test: lua/string_pack/pack_multiple_values_in_order
-- origin: languages/lua/tests/lua/test_string_pack.rs

local __w1 = "1,2,3"
local __i = 0

local s=string.pack("bBi", 1, 2, 3)
local a,b,c=string.unpack("bBi", s)
do local __t = tostring(a..","..b..","..c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
