-- vybe-test: lua/functions_vararg/test_vararg_not_in_vararg_function
-- origin: languages/lua/tests/lua/test_functions_vararg.rs

local __w1 = "true"
local __i = 0

local ok = load('local function f() return ... end'); do local __t = tostring(tostring(ok==nil)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
