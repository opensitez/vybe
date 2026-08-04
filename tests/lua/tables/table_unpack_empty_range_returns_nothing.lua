-- vybe-test: lua/tables/table_unpack_empty_range_returns_nothing
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "nil,nil"
local __i = 0

local a, b = table.unpack({10, 20}, 3, 2)
do local __t = tostring(tostring(a) .. "," .. tostring(b)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
