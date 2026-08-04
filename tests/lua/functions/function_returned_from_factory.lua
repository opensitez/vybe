-- vybe-test: lua/functions/function_returned_from_factory
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "13"
local __i = 0

local function make_adder(n)
  return function(x) return x + n end
end
do local __t = tostring(make_adder(10)(3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
