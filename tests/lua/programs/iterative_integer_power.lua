-- vybe-test: lua/programs/iterative_integer_power
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "32"
local __i = 0

local base, exp = 2, 5
local r = 1
for i = 1, exp do r = r * base end
do local __t = tostring(r); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
