-- vybe-test: lua/debug_library/debug_getinfo_returns_function_metadata
-- origin: languages/lua/tests/lua/test_debug_library.rs

local __w1 = "true"
local __i = 0

local function f() end
local info = debug.getinfo(f)
do local __t = tostring(info.source ~= nil); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
