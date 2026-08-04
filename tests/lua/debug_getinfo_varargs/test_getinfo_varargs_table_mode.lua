-- vybe-test: lua/debug_getinfo_varargs/test_getinfo_varargs_table_mode
-- origin: languages/lua/tests/lua/test_debug_getinfo_varargs.rs

local __w1 = "true"
local __i = 0

local function f(...)
  local info = debug.getinfo(1, "u")
  do local __t = tostring(type(info) == "table"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end
f()

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
