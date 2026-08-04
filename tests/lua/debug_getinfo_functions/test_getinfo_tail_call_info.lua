-- vybe-test: lua/debug_getinfo_functions/test_getinfo_tail_call_info
-- origin: languages/lua/tests/lua/test_debug_getinfo_functions.rs

local __w1 = "Lua"
local __i = 0

local function a()
  return b()
end
function b()
  local info = debug.getinfo(1, "S")
  return info.what
end
do local __t = tostring(a()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
