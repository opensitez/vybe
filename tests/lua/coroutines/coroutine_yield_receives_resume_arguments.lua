-- vybe-test: lua/coroutines/coroutine_yield_receives_resume_arguments
-- origin: languages/lua/tests/lua/test_coroutines.rs

local __w1 = "99"
local __i = 0

local co = coroutine.create(function()
  local x = coroutine.yield()
  return x
end)
coroutine.resume(co)
local _, v = coroutine.resume(co, 99)
do local __t = tostring(v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
