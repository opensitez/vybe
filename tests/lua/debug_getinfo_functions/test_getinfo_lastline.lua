-- vybe-test: lua/debug_getinfo_functions/test_getinfo_lastline
-- origin: languages/lua/tests/lua/test_debug_getinfo_functions.rs

local __w1 = "true"
local __i = 0

local function f() end
local info = debug.getinfo(f, "S")
do local __t = tostring(type(info.lastlinedefined) == "number"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
