-- vybe-test: lua/coroutine_patterns/coroutine_pipeline
-- origin: languages/lua/tests/lua/test_coroutine_patterns.rs

local __w1 = "30"
local __i = 0

local function filter(gen, pred)
  return coroutine.wrap(function()
    for v in gen do
      if pred(v) then coroutine.yield(v) end
    end
  end)
end
local function nums(n)
  return coroutine.wrap(function()
    for i = 1, n do coroutine.yield(i) end
  end)
end
local evens = filter(nums(10), function(v) return v % 2 == 0 end)
local sum = 0
for v in evens do sum = sum + v end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
