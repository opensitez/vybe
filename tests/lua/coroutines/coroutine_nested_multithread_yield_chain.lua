-- vybe-test: lua/coroutines/coroutine_nested_multithread_yield_chain
-- origin: languages/lua/tests/lua/test_coroutines.rs

local __w1 = "main resume A end\ncoB yield end\ncoA resume B end"
local __i = 0

local coB
local coA = coroutine.create(function()
  coroutine.resume(coB)
  do local __t = tostring("coA resume B end"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end)
coB = coroutine.create(function()
  coroutine.yield()
  do local __t = tostring("coB yield end"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end)
coroutine.resume(coA)
do local __t = tostring("main resume A end"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
coroutine.resume(coB)

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
