-- vybe-test: lua/programs/rotate_array_left_by_one
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "3"
local __i = 0

local t = {2, 3, 4, 5}
local first = t[1]
for i = 1, #t - 1 do t[i] = t[i + 1] end
t[#t] = first
do local __t = tostring(t[1]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
