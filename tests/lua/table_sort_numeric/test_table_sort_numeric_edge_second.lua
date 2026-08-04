-- vybe-test: lua/table_sort_numeric/test_table_sort_numeric_edge_second
-- origin: languages/lua/tests/lua/test_table_sort_numeric.rs

local __w1 = "true"
local __i = 0

local t = {}
for i=1,18 do t[i] = 18 - i end
table.sort(t)
do local __t = tostring(t[1] == 0); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
