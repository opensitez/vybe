-- vybe-test: lua/loops_for_generic_pairs_mut/test_loops_for_generic_pairs_mut_sum_with_mutation_inside_loop
-- origin: languages/lua/tests/lua/test_loops_for_generic_pairs_mut.rs

local __w1 = "6"
local __i = 0

local t = {a = 1, b = 2}
local total = 0
for k, v in pairs(t) do t[k] = v * 2; total = total + t[k] end
do local __t = tostring(total); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
