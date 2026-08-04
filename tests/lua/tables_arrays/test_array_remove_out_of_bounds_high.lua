-- vybe-test: lua/tables_arrays/test_array_remove_out_of_bounds_high
-- origin: languages/lua/tests/lua/test_tables_arrays.rs

local __w1 = "false"
local __i = 0

local t={1}; local ok = pcall(function() table.remove(t, 2) end); do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
