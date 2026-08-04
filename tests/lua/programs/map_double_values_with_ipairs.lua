-- vybe-test: lua/programs/map_double_values_with_ipairs
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "6"
local __i = 0

local t = {2, 3, 4}
for i, v in ipairs(t) do t[i] = v * 2 end
do local __t = tostring(t[2]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
