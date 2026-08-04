-- vybe-test: lua/programs/linear_search_found_index
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "2"
local __i = 0

local t = {5, 9, 3}
local target = 9
local idx = 0
for i = 1, #t do if t[i] == target then idx = i break end end
do local __t = tostring(idx); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
