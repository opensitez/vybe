-- vybe-test: lua/coroutines/coroutine_dead_after_error_in_body
-- origin: languages/lua/tests/lua/test_coroutines.rs

local __w1 = "dead"
local __i = 0

local co = coroutine.create(function() error("x") end)
pcall(coroutine.resume, co)
do local __t = tostring(coroutine.status(co)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
