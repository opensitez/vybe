-- vybe-test: lua/debug_locals/debug_getlocal_on_active_coroutine
-- origin: languages/lua/tests/lua/test_debug_locals.rs

local __w1 = "x 42"
local __i = 0

local co = coroutine.create(function(x) local y = 10; coroutine.yield() end)
coroutine.resume(co, 42)
local name, val = debug.getlocal(co, 1, 1)
do local __t = tostring(name .. " " .. val); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
