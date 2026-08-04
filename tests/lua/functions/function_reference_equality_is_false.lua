-- vybe-test: lua/functions/function_reference_equality_is_false
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "true"
local __i = 0

local f = function() end
do local __t = tostring(tostring(f == f)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
