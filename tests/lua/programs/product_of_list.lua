-- vybe-test: lua/programs/product_of_list
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "24"
local __i = 0

local t = {2, 3, 4}
local p = 1
for i = 1, #t do p = p * t[i] end
do local __t = tostring(p); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
