-- vybe-test: lua/coroutines_nested_yield/multiple_nested_yields
-- origin: languages/lua/tests/lua/test_coroutines_nested_yield.rs

local __w1 = "15"
local __i = 0

local function step()
  local x = coroutine.yield("need_x")
  local y = coroutine.yield("need_y")
  return x + y
end
local co = coroutine.create(step)
coroutine.resume(co)
coroutine.resume(co, 5)
local _, sum = coroutine.resume(co, 10)
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
