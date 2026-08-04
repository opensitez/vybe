-- vybe-test: lua/closures/closure_captures_parameter_value
-- origin: languages/lua/tests/lua/test_closures.rs

local __w1 = "3"
local __i = 0

function make(x)
  return function() return x end
end
do local __t = tostring(make(3)()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
