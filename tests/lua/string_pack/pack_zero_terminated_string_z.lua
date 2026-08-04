-- vybe-test: lua/string_pack/pack_zero_terminated_string_z
-- origin: languages/lua/tests/lua/test_string_pack.rs

local __w1 = "end\t5"
local __i = 0

local s=string.pack("z", "end")
do local __t = tostring(string.unpack("z", s)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
