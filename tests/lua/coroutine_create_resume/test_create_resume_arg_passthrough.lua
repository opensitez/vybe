-- vybe-test: lua/coroutine_create_resume/test_create_resume_arg_passthrough
-- origin: languages/lua/tests/lua/test_coroutine_create_resume.rs

local __w1 = "true"
local __i = 0

local t = coroutine.create(function(v) return v end)
local ok, v = coroutine.resume(t, 11)
do local __t = tostring(ok and v == 11); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
