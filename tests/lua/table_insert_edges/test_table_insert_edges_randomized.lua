-- vybe-test: lua/table_insert_edges/test_table_insert_edges_randomized
-- origin: languages/lua/tests/lua/test_table_insert_edges.rs

local __w1 = "false"
local __i = 0

local t = {1,2,3}; local ok = pcall(function() table.insert(t, 5, 4) end); do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
