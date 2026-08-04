-- vybe-test: lua/coroutines/coroutine_running_inside_coroutine_is_thread
-- origin: languages/lua/tests/lua/test_coroutines.rs

local __w1 = "thread"
local __i = 0

local co = coroutine.create(function() do local __t = tostring(type(coroutine.running())); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end)
coroutine.resume(co)

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
