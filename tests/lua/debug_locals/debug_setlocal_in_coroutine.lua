-- vybe-test: lua/debug_locals/debug_setlocal_in_coroutine
-- origin: languages/lua/tests/lua/test_debug_locals.rs

local __w1 = "42"
local __i = 0

local co = coroutine.create(function(x) local y = 10; coroutine.yield() end)
coroutine.resume(co, 42)
debug.setlocal(co, 1, 1, 99)
local _, val = debug.getlocal(co, 1, 1)
do local __t = tostring(val); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
