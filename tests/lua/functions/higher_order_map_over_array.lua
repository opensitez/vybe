-- vybe-test: lua/functions/higher_order_map_over_array
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "9"
local __i = 0

local function map(t, f)
  local out = {}
  for i = 1, #t do out[i] = f(t[i]) end
  return out
end
do local __t = tostring(map({1, 2, 3}, function(x) return x * x end)[3]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
