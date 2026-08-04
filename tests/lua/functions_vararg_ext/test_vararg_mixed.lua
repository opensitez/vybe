-- vybe-test: lua/functions_vararg_ext/test_vararg_mixed
-- origin: languages/lua/tests/lua/test_functions_vararg_ext.rs

local __w1 = "1 2"
local __i = 0

local function f(a, ...) return a, select('#', ...) end; local x, c = f(1, 2, 3); do local __t = tostring(x..' '..c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
