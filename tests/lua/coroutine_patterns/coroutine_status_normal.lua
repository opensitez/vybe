-- vybe-test: lua/coroutine_patterns/coroutine_status_normal
-- origin: languages/lua/tests/lua/test_coroutine_patterns.rs

local __w1 = "running"
local __i = 0

local main_status
local co2 = coroutine.create(function()
  main_status = coroutine.status(coroutine.running())
end)
coroutine.resume(co2)
do local __t = tostring(main_status); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
