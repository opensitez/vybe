-- vybe-test: lua/coroutine_create_resume/test_create_resume_no_function
-- origin: languages/lua/tests/lua/test_coroutine_create_resume.rs

local __w1 = "true"
local __i = 0

local ok, err = pcall(function() coroutine.create(1) end)
do local __t = tostring(ok == false); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
