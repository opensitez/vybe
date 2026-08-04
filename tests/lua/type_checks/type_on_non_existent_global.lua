-- vybe-test: lua/type_checks/type_on_non_existent_global
-- origin: languages/lua/tests/lua/test_type_checks.rs

local __w1 = "nil"
local __i = 0

do local __t = tostring(type(non_existent_global_var_xyz_123)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
