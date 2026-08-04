-- vybe-test: lua/functions/function_recursion_factorial
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "120"
local __i = 0

function fact(n)
  if n <= 1 then return 1 end
  return n * fact(n - 1)
end
do local __t = tostring(fact(5)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
