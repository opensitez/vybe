-- vybe-test: lua/programs/first_matching_predicate_index
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "4"
local __i = 0

local t = {2, 4, 6, 7}
local idx = 0
for i = 1, #t do if t[i] % 2 == 1 then idx = i break end end
do local __t = tostring(idx); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
