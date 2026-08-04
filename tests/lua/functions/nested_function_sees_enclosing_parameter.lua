-- vybe-test: lua/functions/nested_function_sees_enclosing_parameter
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "8"
local __i = 0

function outer(x)
  function inner() return x end
  return inner()
end
do local __t = tostring(outer(8)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
