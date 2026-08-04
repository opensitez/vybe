-- vybe-test: lua/vararg_functions_advanced/vararg_select_rest
-- origin: languages/lua/tests/lua/test_vararg_functions_advanced.rs

local __w1 = "b\tc"
local __i = 0

local function f(...) return select(2, ...) end
do local __t = tostring(f("a", "b", "c")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
