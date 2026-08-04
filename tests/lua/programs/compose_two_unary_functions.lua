-- vybe-test: lua/programs/compose_two_unary_functions
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "7"
local __i = 0

local function compose(f, g) return function(x) return f(g(x)) end end
local inc = function(n) return n + 1 end
local double = function(n) return n * 2 end
do local __t = tostring(compose(inc, double)(3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
