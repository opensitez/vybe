-- vybe-test: lua/iteration/ipairs_does_not_visit_hash_keys
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "5"
local __i = 0

local t = {x = 1, 2, 3}
local sum = 0
for _, v in ipairs(t) do sum = sum + v end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
