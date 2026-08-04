-- vybe-test: lua/programs/find_minimum_in_array
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "1"
local __i = 0

local t = {4, 1, 9, 2}
local min = t[1]
for i = 2, #t do if t[i] < min then min = t[i] end end
do local __t = tostring(min); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
