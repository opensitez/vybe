-- vybe-test: lua/table_super/tbl_sup_insert_6
-- origin: languages/lua/tests/lua/test_table_super.rs

local __w1 = "6"
local __i = 0

local t = {}
table.insert(t, 6)
do local __t = tostring(t[1]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
