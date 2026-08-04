-- vybe-test: lua/coroutine_create_resume/test_create_local_parameter
-- origin: languages/lua/tests/lua/test_coroutine_create_resume.rs

local __w1 = "true"
local __i = 0

local t = coroutine.create(function(a) return a * 2 end)
local ok, v = coroutine.resume(t, 4)
do local __t = tostring(ok and v == 8); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
