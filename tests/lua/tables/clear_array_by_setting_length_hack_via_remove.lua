-- vybe-test: lua/tables/clear_array_by_setting_length_hack_via_remove
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "0"
local __i = 0

local t = {1, 2, 3}
while #t > 0 do table.remove(t) end
do local __t = tostring(#t); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
