-- vybe-test: lua/functions/function_returns_boolean_for_if
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "yes"
local __i = 0

function is_even(n) return n % 2 == 0 end
if is_even(4) then do local __t = tostring("yes"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
