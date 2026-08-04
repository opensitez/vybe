-- vybe-test: lua/functions/immediately_invoked_function_expression
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "5"
local __i = 0

do local __t = tostring((function(x) return x + 1 end)(4)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
