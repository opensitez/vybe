-- vybe-test: lua/programs/parity_even_check
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "true"
local __i = 0

local n = 4
do local __t = tostring(n % 2 == 0); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
