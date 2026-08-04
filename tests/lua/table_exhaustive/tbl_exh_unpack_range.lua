-- vybe-test: lua/table_exhaustive/tbl_exh_unpack_range
-- origin: languages/lua/tests/lua/test_table_exhaustive.rs

local __w1 = "20\t30"
local __i = 0

local a, b = table.unpack({10, 20, 30, 40}, 2, 3)
do local __t = tostring(a) .. "\t" .. tostring(b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
