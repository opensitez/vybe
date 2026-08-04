-- vybe-test: lua/coroutine_resume_values/test_resume_status_running_to_suspended
-- origin: languages/lua/tests/lua/test_coroutine_resume_values.rs

local __w1 = "true"
local __i = 0

local t = coroutine.create(function() coroutine.yield(true) end)
local ok, v = coroutine.resume(t)
do local __t = tostring(ok and v == true and coroutine.status(t) == "suspended"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
