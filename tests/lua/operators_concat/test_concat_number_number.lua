-- vybe-test: lua/operators_concat/test_concat_number_number
-- origin: languages/lua/tests/lua/test_operators_concat.rs

local __w1 = "1020"
local __i = 0

do local __t = tostring(10 .. 20); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
