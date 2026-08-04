-- vybe-test: lua/coroutine_create_resume/test_create_basic_thread
-- origin: languages/lua/tests/lua/test_coroutine_create_resume.rs

local __w1 = "thread"
local __i = 0

local t = coroutine.create(function() return 1 end)
do local __t = tostring(type(t)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
