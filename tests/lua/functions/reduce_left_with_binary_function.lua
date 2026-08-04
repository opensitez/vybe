-- vybe-test: lua/functions/reduce_left_with_binary_function
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "6"
local __i = 0

local function fold(t, f, init)
  local acc = init
  for i = 1, #t do acc = f(acc, t[i]) end
  return acc
end
do local __t = tostring(fold({1, 2, 3}, function(a, b) return a + b end, 0)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
