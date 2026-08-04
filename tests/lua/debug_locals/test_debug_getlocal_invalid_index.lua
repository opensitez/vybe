-- vybe-test: lua/debug_locals/test_debug_getlocal_invalid_index
-- origin: languages/lua/tests/lua/test_debug_locals.rs

local __w1 = "nil nil"
local __i = 0

local n, v = debug.getlocal(1, 100); do local __t = tostring(tostring(n)..' '..tostring(v)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
