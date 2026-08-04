-- vybe-test: lua/functions/function_default_nil_param
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "nil"
local __i = 0

function show(x) return tostring(x) end
do local __t = tostring(show()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
