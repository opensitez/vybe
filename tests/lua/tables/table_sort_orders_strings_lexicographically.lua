-- vybe-test: lua/tables/table_sort_orders_strings_lexicographically
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "a,b,c"
local __i = 0

local t = {"b", "a", "c"}
table.sort(t)
do local __t = tostring(table.concat(t, ",")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
