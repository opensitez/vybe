-- vybe-test: lua/truthiness/type_of_coroutine_is_thread
-- origin: languages/lua/tests/lua/test_truthiness.rs

local __w1 = "thread"
local __i = 0

local co = coroutine.create(function() end)
do local __t = tostring(type(co)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
