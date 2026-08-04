-- vybe-test: lua/coroutine_patterns/coroutine_accumulate
-- origin: languages/lua/tests/lua/test_coroutine_patterns.rs

local __w1 = "30"
local __i = 0

local co = coroutine.create(function()
  for i = 1, 4 do coroutine.yield(i * i) end
end)
local sum = 0
for _ = 1, 4 do
  local _, v = coroutine.resume(co)
  sum = sum + v
end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
