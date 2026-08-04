-- vybe-test: lua/table_concat/test_concat_out_of_bounds
-- origin: languages/lua/tests/lua/test_table_concat.rs

local __w1 = "false"
local __i = 0

local ok, err = pcall(function() table.concat({'a'}, ',', 1, 3) end); do local __t = tostring(tostring(ok)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
