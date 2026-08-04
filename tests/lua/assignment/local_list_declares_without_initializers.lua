-- vybe-test: lua/assignment/local_list_declares_without_initializers
-- origin: languages/lua/tests/lua/test_assignment.rs

local __w1 = "3"
local __i = 0

local a, b
c = 1
b = 2
a = b + c
do local __t = tostring(a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
