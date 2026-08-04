-- vybe-test: lua/iteration/ipairs_ignores_hash_part_of_mixed_table
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "30"
local __i = 0

local t = {10, 20, x = 99}
local sum = 0
for _, v in ipairs(t) do sum = sum + v end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
