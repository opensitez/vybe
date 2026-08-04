-- vybe-test: lua/globals/setting_global_via_g_table_is_readable_by_name
-- origin: languages/lua/tests/lua/test_globals.rs

local __w1 = "77"
local __i = 0

_G.my_global_val = 77
do local __t = tostring(my_global_val); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
