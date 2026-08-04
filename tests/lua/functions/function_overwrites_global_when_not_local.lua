-- vybe-test: lua/functions/function_overwrites_global_when_not_local
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "hi"
local __i = 0

function greet() return "hi" end
function wrap() return greet() end
do local __t = tostring(wrap()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
