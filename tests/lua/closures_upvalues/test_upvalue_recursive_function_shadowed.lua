-- vybe-test: lua/closures_upvalues/test_upvalue_recursive_function_shadowed
-- origin: languages/lua/tests/lua/test_closures_upvalues.rs

local __w1 = "2400"
local __i = 0

local function f(n) if n==0 then return 1 else return n * f(n-1) end end; local g=f; f=function() return 100 end; do local __t = tostring(g(4)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
