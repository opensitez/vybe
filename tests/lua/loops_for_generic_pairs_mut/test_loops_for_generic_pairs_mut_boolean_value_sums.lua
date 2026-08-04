-- vybe-test: lua/loops_for_generic_pairs_mut/test_loops_for_generic_pairs_mut_boolean_value_sums
-- origin: languages/lua/tests/lua/test_loops_for_generic_pairs_mut.rs

local __w1 = "1"
local __i = 0

local t = {a = true, b = true, c = false}
local total = 0
for _, value in pairs(t) do if value then total = total + 1 else total = total - 1 end end
do local __t = tostring(total); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
