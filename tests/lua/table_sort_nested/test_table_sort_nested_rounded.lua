-- vybe-test: lua/table_sort_nested/test_table_sort_nested_rounded
-- origin: languages/lua/tests/lua/test_table_sort_nested.rs

local __w1 = "true"
local __i = 0

local t = {{v=10}, {v=8}, {v=9}}
table.sort(t, function(a,b) return a.v < b.v end)
do local __t = tostring(t[1].v == 8); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
