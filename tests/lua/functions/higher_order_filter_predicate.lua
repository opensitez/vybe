-- vybe-test: lua/functions/higher_order_filter_predicate
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "2"
local __i = 0

local function keep_if(t, pred)
  local out = {}
  for i = 1, #t do if pred(t[i]) then table.insert(out, t[i]) end end
  return out
end
do local __t = tostring(#keep_if({1, 2, 3, 4}, function(n) return n % 2 == 1 end)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
