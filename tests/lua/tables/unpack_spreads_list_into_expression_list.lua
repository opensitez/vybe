-- vybe-test: lua/tables/unpack_spreads_list_into_expression_list
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "6"
local __i = 0

local function sum(a, b, c) return a + b + c end
do local __t = tostring(sum(table.unpack({1, 2, 3}))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
