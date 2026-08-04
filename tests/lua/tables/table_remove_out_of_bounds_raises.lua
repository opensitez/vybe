-- vybe-test: lua/tables/table_remove_out_of_bounds_raises
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "false"
local __i = 0

local t = {10, 20}
local ok = pcall(function() table.remove(t, 5) end)
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
