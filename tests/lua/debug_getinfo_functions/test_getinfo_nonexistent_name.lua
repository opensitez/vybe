-- vybe-test: lua/debug_getinfo_functions/test_getinfo_nonexistent_name
-- origin: languages/lua/tests/lua/test_debug_getinfo_functions.rs

local __w1 = "true"
local __i = 0

local function f(x, y)
  local info = debug.getinfo(1, "n")
  return info.name
end
do local __t = tostring(type(f(1,2)) == "string"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
