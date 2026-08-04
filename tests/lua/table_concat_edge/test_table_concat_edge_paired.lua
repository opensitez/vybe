-- vybe-test: lua/table_concat_edge/test_table_concat_edge_paired
-- origin: languages/lua/tests/lua/test_table_concat_edge.rs

local __w1 = "true"
local __i = 0

local t = {10, 11, 21}
do local __t = tostring(table.concat(t, ",") == ("10,11,21")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
