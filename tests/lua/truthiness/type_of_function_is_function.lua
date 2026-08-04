-- vybe-test: lua/truthiness/type_of_function_is_function
-- origin: languages/lua/tests/lua/test_truthiness.rs

local __w1 = "function"
local __i = 0

local f = function() end
do local __t = tostring(type(f)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
