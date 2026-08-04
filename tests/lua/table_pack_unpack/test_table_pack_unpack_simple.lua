-- vybe-test: lua/table_pack_unpack/test_table_pack_unpack_simple
-- origin: languages/lua/tests/lua/test_table_pack_unpack.rs

local __w1 = "true"
local __i = 0

local t = table.pack(2, 3, 4)
local a, b, c = table.unpack(t)
do local __t = tostring(a == 2 and b == 3 and c == 4); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
