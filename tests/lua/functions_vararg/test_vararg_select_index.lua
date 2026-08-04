-- vybe-test: lua/functions_vararg/test_vararg_select_index
-- origin: languages/lua/tests/lua/test_functions_vararg.rs

local __w1 = "yz"
local __i = 0

local function f(...) return select(2, ...) end; local a,b = f('x','y','z'); do local __t = tostring(a..b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
