-- vybe-test: lua/table_len_boundaries/test_table_len_boundaries_randomized
-- origin: languages/lua/tests/lua/test_table_len_boundaries.rs

local __w1 = "true"
local __i = 0

local t = {1,2,3}
rawset(t, 22, 23)
do local __t = tostring(#t >= 3); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
