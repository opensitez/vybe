-- vybe-test: lua/tables/table_concat_joins_array_part
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "a,b"
local __i = 0

local t = {}
t[1] = "a"
t[2] = "b"
do local __t = tostring(table.concat(t, ",")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
