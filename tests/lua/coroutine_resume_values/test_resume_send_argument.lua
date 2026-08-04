-- vybe-test: lua/coroutine_resume_values/test_resume_send_argument
-- origin: languages/lua/tests/lua/test_coroutine_resume_values.rs

local __w1 = "true"
local __i = 0

local t = coroutine.create(function(v) return v * 3 end)
local ok1, y = coroutine.resume(t, 3)
do local __t = tostring(ok1 and y == 9); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
