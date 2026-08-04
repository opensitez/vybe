-- vybe-test: lua/type_checks/rawequal_same_table_is_true
-- origin: languages/lua/tests/lua/test_type_checks.rs

local __w1 = "true"
local __i = 0

local a={}
do local __t = tostring(rawequal(a,a)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
