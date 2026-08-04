-- vybe-test: lua/functions/function_multiple_return_values
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "2,1"
local __i = 0

function swap(a,b) return b,a end
local x,y=swap(1,2)
do local __t = tostring(x..","..y); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
