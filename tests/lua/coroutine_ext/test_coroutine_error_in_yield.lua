-- vybe-test: lua/coroutine_ext/test_coroutine_error_in_yield
-- origin: languages/lua/tests/lua/test_coroutine_ext.rs

local __w1 = "1"
local __i = 0

local co = coroutine.create(function() pcall(function() coroutine.yield() end) return 1 end); coroutine.resume(co); local ok, r = coroutine.resume(co); do local __t = tostring(r); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
