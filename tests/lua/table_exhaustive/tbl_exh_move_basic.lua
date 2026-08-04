-- vybe-test: lua/table_exhaustive/tbl_exh_move_basic
-- origin: languages/lua/tests/lua/test_table_exhaustive.rs

local __w1 = "10\t20\t30"
local __i = 0

local a = {10, 20, 30}
local b = {}
table.move(a, 1, 3, 1, b)
do local __t = tostring(b[1]) .. "\t" .. tostring(b[2]) .. "\t" .. tostring(b[3]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
