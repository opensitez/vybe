-- vybe-test: lua/tables_arrays/test_array_insert_with_holes
-- origin: languages/lua/tests/lua/test_tables_arrays.rs

local __w1 = "2"
local __i = 0

local t={1, nil, 3}; table.insert(t, 2); do local __t = tostring(t[4]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
