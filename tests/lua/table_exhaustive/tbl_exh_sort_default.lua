-- vybe-test: lua/table_exhaustive/tbl_exh_sort_default
-- origin: languages/lua/tests/lua/test_table_exhaustive.rs

local __w1 = "1\t2\t3"
local __i = 0

local t = {3, 1, 2}
table.sort(t)
do local __t = tostring(t[1]) .. "\t" .. tostring(t[2]) .. "\t" .. tostring(t[3]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
