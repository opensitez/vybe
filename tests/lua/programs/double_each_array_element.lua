-- vybe-test: lua/programs/double_each_array_element
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "2,4,6"
local __i = 0

local t = {1, 2, 3}
for i = 1, #t do t[i] = t[i] * 2 end
do local __t = tostring(table.concat(t, ",")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
