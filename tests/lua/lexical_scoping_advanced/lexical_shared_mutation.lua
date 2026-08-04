-- vybe-test: lua/lexical_scoping_advanced/lexical_shared_mutation
-- origin: languages/lua/tests/lua/test_lexical_scoping_advanced.rs

local __w1 = "1,11,12"
local __i = 0

local x = 0
local f1 = function() x = x + 1; return x end
local f2 = function() x = x + 10; return x end
do local __t = tostring(f1() .. "," .. f2() .. "," .. f1()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
