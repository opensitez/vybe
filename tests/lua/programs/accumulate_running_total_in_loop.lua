-- vybe-test: lua/programs/accumulate_running_total_in_loop
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "6"
local __i = 0

local t = {1, 2, 3}
local sum = 0
for i = 1, #t do sum = sum + t[i] end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
