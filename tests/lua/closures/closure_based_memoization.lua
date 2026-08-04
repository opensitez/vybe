-- vybe-test: lua/closures/closure_based_memoization
-- origin: languages/lua/tests/lua/test_closures.rs

local __w1 = "2"
local __i = 0

local function memoize(f)
  local cache = {}
  return function(n)
    if cache[n] == nil then cache[n] = f(n) end
    return cache[n]
  end
end
local calls = 0
local expensive = memoize(function(n) calls = calls + 1; return n * n end)
expensive(4); expensive(4); expensive(5)
do local __t = tostring(calls); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
