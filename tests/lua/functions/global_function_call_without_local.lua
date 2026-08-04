-- vybe-test: lua/functions/global_function_call_without_local
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "5"
local __i = 0

function inc(x) return x + 1 end
do local __t = tostring(inc(4)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
