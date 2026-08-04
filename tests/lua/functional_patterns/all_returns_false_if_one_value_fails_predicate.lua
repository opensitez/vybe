-- vybe-test: lua/functional_patterns/all_returns_false_if_one_value_fails_predicate
-- origin: languages/lua/tests/lua/test_functional_patterns.rs

local __w1 = "false"
local __i = 0

local function all(t, pred)
  for _, v in ipairs(t) do if not pred(v) then return false end end
  return true
end
do local __t = tostring(tostring(all({2,4,6,7}, function(x) return x % 2 == 0 end))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
