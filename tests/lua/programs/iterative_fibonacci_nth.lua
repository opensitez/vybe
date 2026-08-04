-- vybe-test: lua/programs/iterative_fibonacci_nth
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "8"
local __i = 0

local n = 6
local a, b = 0, 1
for i = 1, n do a, b = b, a + b end
do local __t = tostring(a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
