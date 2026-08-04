-- vybe-test: lua/coroutine_status_initial/test_status_after_argument_resume
-- origin: languages/lua/tests/lua/test_coroutine_status_initial.rs

local __w1 = "true"
local __i = 0

local t = coroutine.create(function(a)
  local b = coroutine.yield(a)
  return b
end)
coroutine.resume(t, 1)
coroutine.resume(t, 2)
do local __t = tostring(coroutine.status(t) == "dead"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
