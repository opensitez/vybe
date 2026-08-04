-- vybe-test: lua/functional_patterns/any_returns_true_if_predicate_matches_at_least_once
-- origin: languages/lua/tests/lua/test_functional_patterns.rs

local __w1 = "true"
local __i = 0

local function any(t, pred)
  for _, v in ipairs(t) do if pred(v) then return true end end
  return false
end
do local __t = tostring(tostring(any({1,3,5,4}, function(x) return x % 2 == 0 end))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
