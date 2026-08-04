-- vybe-test: lua/functions/chained_calls_with_return_values
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "7"
local __i = 0

function twice(x) return x * 2 end
function add1(x) return x + 1 end
do local __t = tostring(add1(twice(3))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
