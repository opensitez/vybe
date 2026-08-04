-- vybe-test: lua/functional_patterns/reduce_func
-- origin: languages/lua/tests/lua/test_functional_patterns.rs

local __w1 = "15"
local __i = 0

local function reduce(t, f, init)
  local acc = init
  for _, v in ipairs(t) do acc = f(acc, v) end
  return acc
end
do local __t = tostring(reduce({1,2,3,4,5}, function(a, b) return a + b end, 0)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
