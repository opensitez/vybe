-- vybe-test: lua/closures_upvalues/test_upvalue_local_recursive_function
-- origin: languages/lua/tests/lua/test_closures_upvalues.rs

local __w1 = "24"
local __i = 0

local function f(n) if n==0 then return 1 else return n * f(n-1) end end; do local __t = tostring(f(4)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
