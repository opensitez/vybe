-- vybe-test: lua/coroutines/coroutine_status_running_during_body
-- origin: languages/lua/tests/lua/test_coroutines.rs

local __w1 = "running"
local __i = 0

local seen = ""
local co = coroutine.create(function()
  seen = coroutine.status(coroutine.running())
end)
coroutine.resume(co)
do local __t = tostring(seen); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
