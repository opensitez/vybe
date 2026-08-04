-- vybe-test: lua/tables/table_concat_with_start_end_range
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "bc"
local __i = 0

local t = {"a", "b", "c"}
do local __t = tostring(table.concat(t, "", 2, 3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
