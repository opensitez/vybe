-- vybe-test: lua/functional_patterns/flip_swaps_function_arguments
-- origin: languages/lua/tests/lua/test_functional_patterns.rs

local __w1 = "7"
local __i = 0

local function flip(f) return function(a, b) return f(b, a) end end
local sub = function(a, b) return a - b end
local rsub = flip(sub)
do local __t = tostring(rsub(3, 10)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
