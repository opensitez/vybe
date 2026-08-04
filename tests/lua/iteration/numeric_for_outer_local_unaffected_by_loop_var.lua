-- vybe-test: lua/iteration/numeric_for_outer_local_unaffected_by_loop_var
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "99"
local __i = 0

local i = 99
for i = 1, 1 do end
do local __t = tostring(i); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
