-- vybe-test: lua/functional_patterns/memoize_func
-- origin: languages/lua/tests/lua/test_functional_patterns.rs

local __w1 = "1"
local __i = 0

local function memoize(f)
  local cache = {}
  return function(n)
    if cache[n] == nil then cache[n] = f(n) end
    return cache[n]
  end
end
local calls = 0
local slow = memoize(function(n) calls = calls + 1; return n * n end)
slow(5); slow(5); slow(5)
do local __t = tostring(calls); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
