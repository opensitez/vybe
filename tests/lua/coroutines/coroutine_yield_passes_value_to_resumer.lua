-- vybe-test: lua/coroutines/coroutine_yield_passes_value_to_resumer
-- origin: languages/lua/tests/lua/test_coroutines.rs

local __w1 = "10,20"
local __i = 0

local co = coroutine.create(function()
  coroutine.yield(10)
  return 20
end)
local _, a = coroutine.resume(co)
local _, b = coroutine.resume(co)
do local __t = tostring(a .. "," .. b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
