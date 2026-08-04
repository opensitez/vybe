-- vybe-test: lua/basics/concatenate_local_numbers_as_strings
-- origin: languages/lua/tests/lua/test_basics.rs

local __w1 = "12"
local __i = 0

local a, b = 1, 2
do local __t = tostring(a .. b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
