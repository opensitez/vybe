-- vybe-test: lua/functions/function_with_three_parameters
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "6"
local __i = 0

function add3(a, b, c) return a + b + c end
do local __t = tostring(add3(1, 2, 3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
