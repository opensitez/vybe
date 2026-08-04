-- vybe-test: lua/loops_for_generic_pairs_mut/test_loops_for_generic_pairs_mut_nested_loop_sum
-- origin: languages/lua/tests/lua/test_loops_for_generic_pairs_mut.rs

local __w1 = "10"
local __i = 0

local matrix = {x = {1,2}, y = {3,4}}
local out = 0
for _, row in pairs(matrix) do for _, value in ipairs(row) do out = out + value end end
do local __t = tostring(out); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
