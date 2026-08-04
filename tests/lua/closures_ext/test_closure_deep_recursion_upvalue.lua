-- vybe-test: lua/closures_ext/test_closure_deep_recursion_upvalue
-- origin: languages/lua/tests/lua/test_closures_ext.rs

local __w1 = "42"
local __i = 0

local function f(n) if n==0 then return function() return 42 end else return f(n-1) end end; do local __t = tostring(f(10)()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
