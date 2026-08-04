-- vybe-test: lua/functions_vararg/test_vararg_function_call_middle
-- origin: languages/lua/tests/lua/test_functions_vararg.rs

local __w1 = "13nil"
local __i = 0

local function g(a,b,c) return a..b..tostring(c) end; local function f(...) return g(..., 3) end; do local __t = tostring(f(1,2)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
