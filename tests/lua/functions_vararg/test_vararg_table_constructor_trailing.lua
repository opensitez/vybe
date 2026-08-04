-- vybe-test: lua/functions_vararg/test_vararg_table_constructor_trailing
-- origin: languages/lua/tests/lua/test_functions_vararg.rs

local __w1 = "102030"
local __i = 0

local function f(...) local t={10, ...}; return t[1]..t[2]..t[3] end; do local __t = tostring(f(20,30)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
