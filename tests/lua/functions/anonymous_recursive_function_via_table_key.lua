-- vybe-test: lua/functions/anonymous_recursive_function_via_table_key
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "13"
local __i = 0

local fib = {}
fib.calc = function(n)
  if n <= 1 then return n end
  return fib.calc(n - 1) + fib.calc(n - 2)
end
do local __t = tostring(fib.calc(7)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
