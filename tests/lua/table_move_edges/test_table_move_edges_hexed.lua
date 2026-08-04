-- vybe-test: lua/table_move_edges/test_table_move_edges_hexed
-- origin: languages/lua/tests/lua/test_table_move_edges.rs

local __w1 = "true"
local __i = 0

local src = {1,2,3,4,5}
local dst = {}
table.move(src, 2, 2, 1, dst)
do local __t = tostring(type(dst[1]) == "number"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
