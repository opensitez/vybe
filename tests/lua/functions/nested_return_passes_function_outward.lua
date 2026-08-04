-- vybe-test: lua/functions/nested_return_passes_function_outward
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "inner"
local __i = 0

local function outer()
  return function() return "inner" end
end
do local __t = tostring(outer()()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
