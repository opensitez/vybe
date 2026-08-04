-- vybe-test: lua/debug_getinfo_functions/test_getinfo_f_line
-- origin: languages/lua/tests/lua/test_debug_getinfo_functions.rs

local __w1 = "true"
local __i = 0

local function f()
  return debug.getinfo(1, "nS")
end
local info = f()
do local __t = tostring(info.what == "Lua"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
