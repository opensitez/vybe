-- vybe-test: lua/coroutine_resume_values/test_resume_value_function
-- origin: languages/lua/tests/lua/test_coroutine_resume_values.rs

local __w1 = "true"
local __i = 0

local t = coroutine.create(function() return function() end end)
local ok, v = coroutine.resume(t)
do local __t = tostring(ok and type(v) == "function"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
