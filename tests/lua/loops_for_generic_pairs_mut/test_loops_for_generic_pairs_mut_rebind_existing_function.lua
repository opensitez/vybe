-- vybe-test: lua/loops_for_generic_pairs_mut/test_loops_for_generic_pairs_mut_rebind_existing_function
-- origin: languages/lua/tests/lua/test_loops_for_generic_pairs_mut.rs

local __w1 = "3"
local __i = 0

local t = {f = function(a) return a + 1 end}
local seen = 0
for key, value in pairs(t) do if type(value) == 'function' then seen = seen + value(2) end end
do local __t = tostring(seen); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
