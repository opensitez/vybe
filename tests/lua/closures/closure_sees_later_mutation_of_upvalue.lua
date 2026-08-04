-- vybe-test: lua/closures/closure_sees_later_mutation_of_upvalue
-- origin: languages/lua/tests/lua/test_closures.rs

local __w1 = "2"
local __i = 0

local n=1
local f=function() return n end
n=2
do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
