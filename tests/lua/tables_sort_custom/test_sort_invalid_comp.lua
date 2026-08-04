-- vybe-test: lua/tables_sort_custom/test_sort_invalid_comp
-- origin: languages/lua/tests/lua/test_tables_sort_custom.rs

local __w1 = "false"
local __i = 0

local t={1,2}; local ok = pcall(function() table.sort(t, 'not a function') end); do local __t = tostring(tostring(ok)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
