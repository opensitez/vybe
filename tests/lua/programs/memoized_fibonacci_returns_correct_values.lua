-- vybe-test: lua/programs/memoized_fibonacci_returns_correct_values
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "55"
local __i = 0

local cache = {}
local function fib(n)
  if cache[n] then return cache[n] end
  if n <= 1 then return n end
  cache[n] = fib(n-1) + fib(n-2)
  return cache[n]
end
do local __t = tostring(fib(10)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
