-- vybe-test: lua/operators_logical/test_and_true_false
-- origin: languages/lua/tests/lua/test_operators_logical.rs

local __w1 = "false"
local __i = 0

do local __t = tostring(tostring(true and false)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
