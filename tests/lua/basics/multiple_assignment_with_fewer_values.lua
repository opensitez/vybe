-- vybe-test: lua/basics/multiple_assignment_with_fewer_values
-- origin: languages/lua/tests/lua/test_basics.rs

local __w1 = "1,2,nil"
local __i = 0

local a, b, c = 1, 2
do local __t = tostring(tostring(a) .. "," .. tostring(b) .. "," .. tostring(c)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
