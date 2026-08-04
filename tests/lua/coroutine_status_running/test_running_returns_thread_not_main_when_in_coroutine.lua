-- vybe-test: lua/coroutine_status_running/test_running_returns_thread_not_main_when_in_coroutine
-- origin: languages/lua/tests/lua/test_coroutine_status_running.rs

local __w1 = "true"
local __i = 0

local stored
local t = coroutine.create(function()
  local th, main = coroutine.running()
  stored = (main == false)
end)
coroutine.resume(t)
do local __t = tostring(stored == true); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
