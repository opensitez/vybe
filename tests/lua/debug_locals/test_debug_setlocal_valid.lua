-- vybe-test: lua/debug_locals/test_debug_setlocal_valid
-- origin: languages/lua/tests/lua/test_debug_locals.rs

local __w1 = "99"
local __i = 0

local function f() local a=1; debug.setlocal(1, 1, 99); return a end; do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
