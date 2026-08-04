-- vybe-test: lua/functions/function_expression_stored_in_local
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "4"
local __i = 0

local f = function(x) return x - 1 end
do local __t = tostring(f(5)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
