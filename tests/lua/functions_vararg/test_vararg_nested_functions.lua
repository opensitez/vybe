-- vybe-test: lua/functions_vararg/test_vararg_nested_functions
-- origin: languages/lua/tests/lua/test_functions_vararg.rs

local __w1 = "123"
local __i = 0

local function outer(...) return function() return ... end end; local f = outer(1,2,3); local a,b,c = f(); do local __t = tostring(a..b..c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
