-- vybe-test: lua/functions_vararg/test_vararg_table_constructor_middle
-- origin: languages/lua/tests/lua/test_functions_vararg.rs

local __w1 = "1030"
local __i = 0

local function f(...) local t={..., 30}; return t[1]..t[2] end; do local __t = tostring(f(10, 20)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
