-- vybe-test: lua/functions/function_with_if_inside
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "6"
local __i = 0

function abs(n)
  if n < 0 then return -n end
  return n
end
do local __t = tostring(abs(-6)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
