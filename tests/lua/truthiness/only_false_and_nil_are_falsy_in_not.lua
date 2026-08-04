-- vybe-test: lua/truthiness/only_false_and_nil_are_falsy_in_not
-- origin: languages/lua/tests/lua/test_truthiness.rs

local __w1 = "true"
local __i = 0

do local __t = tostring((not false) and (not nil) and (not not 0) and (not not "")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
