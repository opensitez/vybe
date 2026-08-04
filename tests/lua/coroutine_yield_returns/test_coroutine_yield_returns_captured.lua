-- vybe-test: lua/coroutine_yield_returns/test_coroutine_yield_returns_captured
-- origin: languages/lua/tests/lua/test_coroutine_yield_returns.rs

local __w1 = "true"
local __i = 0

local co = coroutine.create(function(x)
  coroutine.yield(x)
  return x + 1, x + 2
end)
local ok1, first = coroutine.resume(co, 15)
local ok2, a, b = coroutine.resume(co)
do local __t = tostring(ok1 and first == 15 and ok2 and a == 16 and b == 17); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
