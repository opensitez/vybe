-- vybe-test: lua/loops_for_generic_ipairs_mut/test_loops_for_generic_ipairs_mut_mutate_in_place
-- origin: languages/lua/tests/lua/test_loops_for_generic_ipairs_mut.rs

local __w1 = "12"
local __i = 0

local t = {1,2,3}
for i, value in ipairs(t) do t[i] = value * 2 end
do local __t = tostring(t[1] + t[2] + t[3]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
