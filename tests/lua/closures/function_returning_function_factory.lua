-- vybe-test: lua/closures/function_returning_function_factory
-- origin: languages/lua/tests/lua/test_closures.rs

local __w1 = "8"
local __i = 0

local function make_adder(n)
  return function(x) return x + n end
end
do local __t = tostring(make_adder(5)(3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
