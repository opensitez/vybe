-- vybe-test: lua/table_remove_edges/test_table_remove_edges_trimmed
-- origin: languages/lua/tests/lua/test_table_remove_edges.rs

local __w1 = "true"
local __i = 0

local t = {3, 4, 5}
local v = table.remove(t, 1)
do local __t = tostring(type(v) == "number"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
