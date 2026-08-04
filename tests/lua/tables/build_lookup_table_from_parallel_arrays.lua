-- vybe-test: lua/tables/build_lookup_table_from_parallel_arrays
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "2"
local __i = 0

local keys = {"a", "b"}
local vals = {1, 2}
local map = {}
for i = 1, #keys do map[keys[i]] = vals[i] end
do local __t = tostring(map.b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
