-- vybe-test: lua/functional_patterns/curry_func
-- origin: languages/lua/tests/lua/test_functional_patterns.rs

local __w1 = "8,15"
local __i = 0

local function curry(f) return function(a) return function(b) return f(a, b) end end end
local add = curry(function(a, b) return a + b end)
local add5 = add(5)
do local __t = tostring(add5(3) .. "," .. add5(10)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
