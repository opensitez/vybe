-- vybe-test: lua/coroutine_create_resume/test_create_with_upvalue_capture
-- origin: languages/lua/tests/lua/test_coroutine_create_resume.rs

local __w1 = "true"
local __i = 0

local x = 7
local t = coroutine.create(function() return x end)
local ok, v = coroutine.resume(t)
do local __t = tostring(ok and v == 7); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
