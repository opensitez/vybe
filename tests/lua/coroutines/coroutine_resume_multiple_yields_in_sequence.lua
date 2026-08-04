-- vybe-test: lua/coroutines/coroutine_resume_multiple_yields_in_sequence
-- origin: languages/lua/tests/lua/test_coroutines.rs

local __w1 = "2"
local __i = 0

local co = coroutine.create(function()
  coroutine.yield(1)
  coroutine.yield(2)
end)
coroutine.resume(co)
local _, a = coroutine.resume(co)
do local __t = tostring(a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
