-- vybe-test: lua/tables_arrays/test_array_unpack_range
-- origin: languages/lua/tests/lua/test_tables_arrays.rs

local __w1 = "234"
local __i = 0

local t={1,2,3,4,5}; local a,b,c = table.unpack(t, 2, 4); do local __t = tostring(a..b..c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
