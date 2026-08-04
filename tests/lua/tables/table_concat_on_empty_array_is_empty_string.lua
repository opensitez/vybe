-- vybe-test: lua/tables/table_concat_on_empty_array_is_empty_string
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = ""
local __i = 0

do local __t = tostring(table.concat({})); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
