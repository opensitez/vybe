-- vybe-test: lua/functions/function_single_param
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "22"
local __i = 0

function double(x) return x * 2 end
do local __t = tostring(double(11)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
