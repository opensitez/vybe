-- vybe-test: lua/functions/mutual_recursion_even_odd
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "true"
local __i = 0

local is_even, is_odd
function is_even(n)
  if n == 0 then return true end
  return is_odd(n - 1)
end
function is_odd(n)
  if n == 0 then return false end
  return is_even(n - 1)
end
do local __t = tostring(is_even(4)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
