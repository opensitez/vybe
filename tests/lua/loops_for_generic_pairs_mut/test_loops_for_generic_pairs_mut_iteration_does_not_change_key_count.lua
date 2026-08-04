-- vybe-test: lua/loops_for_generic_pairs_mut/test_loops_for_generic_pairs_mut_iteration_does_not_change_key_count
-- origin: languages/lua/tests/lua/test_loops_for_generic_pairs_mut.rs

local __w1 = "1"
local __i = 0

local t = {a = 1, b = 2}
local before = 0
for _ in pairs(t) do before = before + 1 end
for k, v in pairs(t) do t[k] = v + 1 end
local after = 0
for _ in pairs(t) do after = after + 1 end
do local __t = tostring(before == after and 1 or 0); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
