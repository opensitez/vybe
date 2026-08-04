-- vybe-test: lua/loops_for_generic_pairs_mut/test_loops_for_generic_pairs_mut_pairs_after_setting_defaults
-- origin: languages/lua/tests/lua/test_loops_for_generic_pairs_mut.rs

local __w1 = "1"
local __i = 0

local t = {x = 10}
for key in pairs(t) do t[key .. '_done'] = true end
do local __t = tostring((t.x_done == nil) and 0 or 1); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
