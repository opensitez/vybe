-- vybe-test: lua/debug_getinfo_varargs/test_getinfo_varargs_boolean_param
-- origin: languages/lua/tests/lua/test_debug_getinfo_varargs.rs

local __w1 = "true"
local __i = 0

local function f(flag, ...)
  return debug.getinfo(1, "u").isvararg
end
do local __t = tostring(f(true) == true); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
