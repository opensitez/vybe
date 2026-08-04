-- vybe-test: lua/debug_hooks/test_debug_gethook_empty
-- origin: languages/lua/tests/lua/test_debug_hooks.rs

local __w1 = "nil  0"
local __i = 0

local h, m, c = debug.gethook(); do local __t = tostring(tostring(h)..' '..tostring(m)..' '..tostring(c)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
