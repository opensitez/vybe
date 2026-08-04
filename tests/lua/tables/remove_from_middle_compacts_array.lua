-- vybe-test: lua/tables/remove_from_middle_compacts_array
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "13"
local __i = 0

local t = {1, 2, 3}
table.remove(t, 2)
do local __t = tostring(table.concat(t, "")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
