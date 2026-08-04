-- vybe-test: lua/tables_arrays/test_array_concat_invalid_type
-- origin: languages/lua/tests/lua/test_tables_arrays.rs

local __w1 = "false"
local __i = 0

local t={'a', true, 'c'}; local ok = pcall(function() table.concat(t) end); do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
