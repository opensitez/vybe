-- vybe-test: lua/tables/pack_stores_varargs_in_table_with_n_field
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "3"
local __i = 0

local p = table.pack(1, 2, 3)
do local __t = tostring(p.n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
