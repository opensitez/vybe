-- vybe-test: lua/loops_for_generic_pairs_mut/test_loops_for_generic_pairs_mut_update_nested_tables
-- origin: languages/lua/tests/lua/test_loops_for_generic_pairs_mut.rs

local __w1 = "5"
local __i = 0

local t = {a = {x = 1}, b = {x = 2}}
for _, v in pairs(t) do v.x = v.x + 1 end
do local __t = tostring(t.a.x + t.b.x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
