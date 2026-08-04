-- vybe-test: lua/iteration/pairs_sum_hash_values
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "5"
local __i = 0

local t = {a = 2, b = 3}
local s = 0
for _, v in pairs(t) do s = s + v end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
