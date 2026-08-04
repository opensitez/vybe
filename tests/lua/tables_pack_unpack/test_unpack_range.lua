-- vybe-test: lua/tables_pack_unpack/test_unpack_range
-- origin: languages/lua/tests/lua/test_tables_pack_unpack.rs

local __w1 = "2 3"
local __i = 0

local a, b = table.unpack({1, 2, 3, 4}, 2, 3); do local __t = tostring(a..' '..b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
