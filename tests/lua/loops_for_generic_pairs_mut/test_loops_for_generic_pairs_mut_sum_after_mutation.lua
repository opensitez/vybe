-- vybe-test: lua/loops_for_generic_pairs_mut/test_loops_for_generic_pairs_mut_sum_after_mutation
-- origin: languages/lua/tests/lua/test_loops_for_generic_pairs_mut.rs

local __w1 = "6"
local __i = 0

local t = {a = 1, b = 2, c = 3}
for _, value in pairs(t) do value = value + 5 end
do local __t = tostring(t.a + t.b + t.c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
