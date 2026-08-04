-- vybe-test: lua/functions_vararg_ext/test_vararg_in_table
-- origin: languages/lua/tests/lua/test_functions_vararg_ext.rs

local __w1 = "ab"
local __i = 0

local function f(...) local t={...}; return t[1]..t[2] end; do local __t = tostring(f('a', 'b')); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
