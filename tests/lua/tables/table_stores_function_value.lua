-- vybe-test: lua/tables/table_stores_function_value
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "5"
local __i = 0

local t = { f = function(x) return x + 1 end }
do local __t = tostring(t.f(4)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
