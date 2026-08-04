-- vybe-test: lua/loops_for_generic_ipairs_mut/test_loops_for_generic_ipairs_mut_assign_none
-- origin: languages/lua/tests/lua/test_loops_for_generic_ipairs_mut.rs

local __w1 = "1"
local __i = 0

local seen = 0
for _, value in ipairs({1, nil, 3}) do if value then seen = seen + 1 end end
do local __t = tostring(seen); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
