-- vybe-test: lua/functions_vararg/test_vararg_select_index_out_of_bounds
-- origin: languages/lua/tests/lua/test_functions_vararg.rs

local __w1 = "nil"
local __i = 0

local function f(...) return select(5, ...) end; local a = f(1,2); do local __t = tostring(a or 'nil'); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
