-- vybe-test: lua/functional_patterns/compose_func
-- origin: languages/lua/tests/lua/test_functional_patterns.rs

local __w1 = "11"
local __i = 0

local function compose(f, g) return function(x) return f(g(x)) end end
local double = function(x) return x * 2 end
local inc = function(x) return x + 1 end
local double_then_inc = compose(inc, double)
do local __t = tostring(double_then_inc(5)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
