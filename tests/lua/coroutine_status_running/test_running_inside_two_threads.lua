-- vybe-test: lua/coroutine_status_running/test_running_inside_two_threads
-- origin: languages/lua/tests/lua/test_coroutine_status_running.rs

local __w1 = "ok"
local __i = 0

local inner_state
local t1 = coroutine.create(function()
  inner_state = coroutine.running() ~= nil
end)
local t2 = coroutine.create(function() coroutine.resume(t1) end)
coroutine.resume(t2)
do local __t = tostring(inner_state and "ok" or "no"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
