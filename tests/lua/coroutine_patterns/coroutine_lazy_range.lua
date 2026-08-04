-- vybe-test: lua/coroutine_patterns/coroutine_lazy_range
-- origin: languages/lua/tests/lua/test_coroutine_patterns.rs

local __w1 = "3"
local __i = 0

local function range(n)
  return coroutine.wrap(function()
    for i = 1, n do coroutine.yield(i) end
  end)
end
local t = {}
for v in range(5) do t[#t+1] = v end
do local __t = tostring(t[3]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
