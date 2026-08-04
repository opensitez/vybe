-- vybe-test: lua/functions/callback_accumulates_in_outer_local
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "6"
local __i = 0

local sum = 0
local function add_each(t, fn)
  for i = 1, #t do sum = sum + fn(t[i]) end
end
add_each({1, 2, 3}, function(x) return x end)
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
