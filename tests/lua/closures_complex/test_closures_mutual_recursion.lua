-- vybe-test: lua/closures_complex/test_closures_mutual_recursion
-- origin: languages/lua/tests/lua/test_closures_complex.rs

local __w1 = "true false"
local __i = 0

local is_even, is_odd;
is_even = function(n)
    if n == 0 then return true else return is_odd(n - 1) end
end
is_odd = function(n)
    if n == 0 then return false else return is_even(n - 1) end
end
do local __t = tostring(tostring(is_even(10)) .. ' ' .. tostring(is_odd(10))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
