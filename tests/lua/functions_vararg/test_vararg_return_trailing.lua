-- vybe-test: lua/functions_vararg/test_vararg_return_trailing
-- origin: languages/lua/tests/lua/test_functions_vararg.rs

local __w1 = "012"
local __i = 0

local function f(...) return 0, ... end; local a,b,c = f(1,2); do local __t = tostring(a..b..c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
