-- vybe-test: lua/programs/lcm_from_gcd_formula
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "12.0"
local __i = 0

local a, b = 4, 6
local x, y = a, b
while y ~= 0 do x, y = y, x % y end
do local __t = tostring(a * b / x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
