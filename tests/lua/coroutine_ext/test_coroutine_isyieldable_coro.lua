-- vybe-test: lua/coroutine_ext/test_coroutine_isyieldable_coro
-- origin: languages/lua/tests/lua/test_coroutine_ext.rs

local __w1 = "true"
local __i = 0

local co = coroutine.create(function() do local __t = tostring(tostring(coroutine.isyieldable())); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end); coroutine.resume(co)

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
