-- vybe-test: lua/type_checks/next_with_nil_starts_iteration
-- origin: languages/lua/tests/lua/test_type_checks.rs

local __w1 = "x"
local __i = 0

local t={x=1}
local k=next(t,nil)
do local __t = tostring(k); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
