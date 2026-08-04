-- vybe-test: lua/coroutine_resume_values/test_resume_yield_value
-- origin: languages/lua/tests/lua/test_coroutine_resume_values.rs

local __w1 = "true"
local __i = 0

local t = coroutine.create(function() coroutine.yield(5); return 8 end)
local ok1, v1 = coroutine.resume(t)
local ok2, v2 = coroutine.resume(t)
do local __t = tostring(ok1 and v1 == 5 and ok2 and v2 == 8); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
