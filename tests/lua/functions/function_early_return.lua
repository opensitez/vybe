-- vybe-test: lua/functions/function_early_return
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "-1"
local __i = 0

function sign(n)
  if n < 0 then return -1 end
  if n > 0 then return 1 end
  return 0
end
do local __t = tostring(sign(-3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
