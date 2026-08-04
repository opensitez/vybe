-- vybe-test: lua/coroutines/coroutine_yield_and_resume_multiple_values
-- origin: languages/lua/tests/lua/test_coroutines.rs

local __w1 = "2,3 20,30"
local __i = 0

local co = coroutine.create(function(x, y)
  local a, b = coroutine.yield(x + 1, y + 1)
  return a + 10, b + 10
end)
local _, r1, r2 = coroutine.resume(co, 1, 2)
local _, r3, r4 = coroutine.resume(co, 10, 20)
do local __t = tostring(r1 .. "," .. r2 .. " " .. r3 .. "," .. r4); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
