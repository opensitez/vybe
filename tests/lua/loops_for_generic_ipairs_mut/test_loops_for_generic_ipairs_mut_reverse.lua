-- vybe-test: lua/loops_for_generic_ipairs_mut/test_loops_for_generic_ipairs_mut_reverse
-- origin: languages/lua/tests/lua/test_loops_for_generic_ipairs_mut.rs

local __w1 = "0"
local __i = 0

local t = {1,2,3}
local out = 0
for i, value in ipairs(t) do out = out + value end
t = {3,2,1}
do local __t = tostring((t[1] + out) == 7 and 1 or 0); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
