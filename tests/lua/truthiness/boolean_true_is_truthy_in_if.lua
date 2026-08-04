-- vybe-test: lua/truthiness/boolean_true_is_truthy_in_if
-- origin: languages/lua/tests/lua/test_truthiness.rs

local __w1 = "yes"
local __i = 0

if true then do local __t = tostring("yes"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
