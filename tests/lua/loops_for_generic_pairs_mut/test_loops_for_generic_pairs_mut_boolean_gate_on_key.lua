-- vybe-test: lua/loops_for_generic_pairs_mut/test_loops_for_generic_pairs_mut_boolean_gate_on_key
-- origin: languages/lua/tests/lua/test_loops_for_generic_pairs_mut.rs

local __w1 = "2"
local __i = 0

local t = {one = 1, two = 2, three = 3}
local hits = 0
for key, value in pairs(t) do if key == 'two' then hits = hits + value end end
do local __t = tostring(hits); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
