-- vybe-test: lua/table_super/tbl_sup_remove_3
-- origin: languages/lua/tests/lua/test_table_super.rs

local __w1 = "2"
local __i = 0

local t = {1, 2, 3}
table.remove(t)
do local __t = tostring(#t); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
