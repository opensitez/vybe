-- vybe-test: lua/table_super/tbl_sup_remove_9
-- origin: languages/lua/tests/lua/test_table_super.rs

local __w1 = "8"
local __i = 0

local t = {1, 2, 3, 4, 5, 6, 7, 8, 9}
table.remove(t)
do local __t = tostring(#t); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
