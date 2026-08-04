-- vybe-test: lua/coroutine_create_resume/test_create_yielding_returns_value
-- origin: languages/lua/tests/lua/test_coroutine_create_resume.rs

local __w1 = "true"
local __i = 0

local t = coroutine.create(function() coroutine.yield(9); return 3 end)
coroutine.resume(t)
local ok, v = coroutine.resume(t)
do local __t = tostring(ok and v == 3); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
