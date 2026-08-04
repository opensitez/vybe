-- vybe-test: lua/tables/table_concat_invalid_range_returns_empty
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = ""
local __i = 0

local t = {10, 20, 30}
do local __t = tostring(table.concat(t, ",", 3, 2)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
