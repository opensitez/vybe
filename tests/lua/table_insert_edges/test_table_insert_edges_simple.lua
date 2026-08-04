-- vybe-test: lua/table_insert_edges/test_table_insert_edges_simple
-- origin: languages/lua/tests/lua/test_table_insert_edges.rs

local __w1 = "true"
local __i = 0

local t = {1,2,3}; table.insert(t, 3, 2); do local __t = tostring(t[3] == 2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
