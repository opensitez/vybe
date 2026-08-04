-- vybe-test: lua/coroutines/coroutine_status_normal_during_resume_nested
-- origin: languages/lua/tests/lua/test_coroutines.rs

local __w1 = "normal"
local __i = 0

local co1, co2
co1 = coroutine.create(function()
  coroutine.resume(co2)
end)
co2 = coroutine.create(function()
  do local __t = tostring(coroutine.status(co1)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end)
coroutine.resume(co1)

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
