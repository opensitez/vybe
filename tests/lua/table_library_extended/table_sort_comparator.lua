-- vybe-test: lua/table_library_extended/table_sort_comparator
-- origin: languages/lua/tests/lua/test_table_library_extended.rs

local __w1 = "3,1"
local __i = 0

local t = {1, 3, 2}
table.sort(t, function(a, b) return a > b end)
do local __t = tostring(t[1] .. "," .. t[3]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
