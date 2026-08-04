-- vybe-test: lua/coroutine_ext/test_coroutine_running_coro
-- origin: languages/lua/tests/lua/test_coroutine_ext.rs

local __w1 = "thread false"
local __i = 0

local co = coroutine.create(function() local c, m = coroutine.running(); do local __t = tostring(type(c)..' '..tostring(m)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end); coroutine.resume(co)

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
