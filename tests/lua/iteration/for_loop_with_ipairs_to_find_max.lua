-- vybe-test: lua/iteration/for_loop_with_ipairs_to_find_max
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "9"
local __i = 0

local t = {3, 9, 1}
local max = t[1]
for _, v in ipairs(t) do if v > max then max = v end end
do local __t = tostring(max); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
