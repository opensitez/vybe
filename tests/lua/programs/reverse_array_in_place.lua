-- vybe-test: lua/programs/reverse_array_in_place
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "4,3,2,1"
local __i = 0

local t = {1, 2, 3, 4}
local i, j = 1, #t
while i < j do t[i], t[j] = t[j], t[i] i = i + 1 j = j - 1 end
do local __t = tostring(table.concat(t, ",")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
