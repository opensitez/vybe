-- vybe-test: lua/truthiness/nil_not_equal_to_false_in_equality_check
-- origin: languages/lua/tests/lua/test_truthiness.rs

local __w1 = "false"
local __i = 0

do local __t = tostring(nil == false); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
