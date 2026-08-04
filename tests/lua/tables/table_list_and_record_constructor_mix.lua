-- vybe-test: lua/tables/table_list_and_record_constructor_mix
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "15"
local __i = 0

local t = {10, a = 5}
do local __t = tostring(t[1] + t.a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
