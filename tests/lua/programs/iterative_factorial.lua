-- vybe-test: lua/programs/iterative_factorial
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "120"
local __i = 0

local n = 5
local acc = 1
for i = 2, n do acc = acc * i end
do local __t = tostring(acc); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
