-- vybe-test: lua/loops_for_generic_pairs_mut/test_loops_for_generic_pairs_mut_sum_of_lengthy_keys
-- origin: languages/lua/tests/lua/test_loops_for_generic_pairs_mut.rs

local __w1 = "12"
local __i = 0

local t = {left = 4, right = 8}
local sum = 0
for key, value in pairs(t) do if type(key) == "string" then sum = sum + value end end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
