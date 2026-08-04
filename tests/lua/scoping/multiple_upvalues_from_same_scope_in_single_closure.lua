-- vybe-test: lua/scoping/multiple_upvalues_from_same_scope_in_single_closure
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "6"
local __i = 0

local a, b, c = 1, 2, 3
local f = function() return a + b + c end
do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
