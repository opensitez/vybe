-- vybe-test: lua/functions_vararg/test_vararg_nested_varargs_shadowing
-- origin: languages/lua/tests/lua/test_functions_vararg.rs

local __w1 = "12"
local __i = 0

local function outer(...) local a = ...; return function(...) local b = ...; return a..b end end; do local __t = tostring(outer(1)(2)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
