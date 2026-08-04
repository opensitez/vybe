-- vybe-test: lua/closures/recursive_closure_via_upvalue_self_reference
-- origin: languages/lua/tests/lua/test_closures.rs

local __w1 = "120"
local __i = 0

local fact
fact = function(n)
  if n <= 1 then return 1 end
  return n * fact(n - 1)
end
do local __t = tostring(fact(5)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
