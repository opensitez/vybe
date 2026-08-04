-- vybe-test: lua/types_exhaustive/type_thread_check
-- origin: languages/lua/tests/lua/test_types_exhaustive.rs

local __w1 = "thread"
local __i = 0

do local __t = tostring(type(coroutine.create(function() end))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
