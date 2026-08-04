-- vybe-test: lua/vararg_functions_advanced/vararg_forward_print
-- origin: languages/lua/tests/lua/test_vararg_functions_advanced.rs

local __w1 = "1\t2"
local __i = 0

local function f(...) do local __t = tostring(...); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end
f(1, 2)

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
