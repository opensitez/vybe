-- vybe-test: lua/loops_for_generic_pairs_mut/test_loops_for_generic_pairs_mut_key_projection
-- origin: languages/lua/tests/lua/test_loops_for_generic_pairs_mut.rs

local __w1 = "1"
local __i = 0

local t = {alpha = 1, beta = 2}
local out = ''
for k, value in pairs(t) do out = out .. k end
do local __t = tostring(#out == 9 and 1 or 0); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
