-- vybe-test: lua/debug_getinfo_varargs/test_getinfo_varargs_multiple_calls
-- origin: languages/lua/tests/lua/test_debug_getinfo_varargs.rs

local __w1 = "true"
local __i = 0

local function f(a, ...)
  local info = debug.getinfo(1, "u")
  return info.nparams == 1 and info.isvararg == true
end
do local __t = tostring(f(1,2,3,4)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
