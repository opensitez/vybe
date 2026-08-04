-- vybe-test: lua/functions/unary_function_composed_inline
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "30"
local __i = 0

local function pipe(x, f, g) return g(f(x)) end
do local __t = tostring(pipe(2, function(n) return n + 1 end, function(n) return n * 10 end)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
