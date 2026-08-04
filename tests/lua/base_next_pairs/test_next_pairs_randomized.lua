-- vybe-test: lua/base_next_pairs/test_next_pairs_randomized
-- origin: languages/lua/tests/lua/test_base_next_pairs.rs

local __w1 = "true"
local __i = 0

local t = {}
for i = 1, 19 do t[i] = i end
local c = 0
for k in next, t, nil do c = c + 1 end
do local __t = tostring(c >= 1); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
