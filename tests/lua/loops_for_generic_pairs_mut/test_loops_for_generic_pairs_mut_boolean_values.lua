-- vybe-test: lua/loops_for_generic_pairs_mut/test_loops_for_generic_pairs_mut_boolean_values
-- origin: languages/lua/tests/lua/test_loops_for_generic_pairs_mut.rs

local __w1 = "2"
local __i = 0

local t = {a = true, b = false, c = true}
local sum = 0
for _, value in pairs(t) do if value then sum = sum + 1 end end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
