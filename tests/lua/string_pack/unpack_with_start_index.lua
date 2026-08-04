-- vybe-test: lua/string_pack/unpack_with_start_index
-- origin: languages/lua/tests/lua/test_string_pack.rs

local __w1 = "1\t2"
local __i = 0

local s=string.pack("bb", 1, 2)
do local __t = tostring(string.unpack("b", s)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
