-- vybe-test: lua/table_exhaustive/tbl_exh_remove_pos
-- origin: languages/lua/tests/lua/test_table_exhaustive.rs

local __w1 = "20\t10\t30"
local __i = 0

local t = {10, 20, 30}
local v = table.remove(t, 2)
do local __t = tostring(v) .. "\t" .. tostring(t[1]) .. "\t" .. tostring(t[2]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
