-- 16_memoization_closure.lua
-- Demonstrates closure-based memoization for expensive recursive functions.

local function memoize(fn)
  local cache = {}
  return function(n)
    if cache[n] ~= nil then
      return cache[n], true
    end
    local value = fn(n)
    cache[n] = value
    return value, false
  end
end

local fib
fib = memoize(function(n)
  if n < 2 then
    return n
  end
  local a = fib(n - 1)
  local b = fib(n - 2)
  return a + b
end)

for i = 0, 20 do
  local value, from_cache = fib(i)
  print(i, value, from_cache and "cache" or "computed")
end

local _, cached = fib(20)
print("fib(20) second call cache hit:", cached)
