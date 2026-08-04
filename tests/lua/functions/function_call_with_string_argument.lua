-- vybe-test: lua/functions/function_call_with_string_argument
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "HI"
local __i = 0

function shout(s) return string.upper(s) end
do local __t = tostring(shout("hi")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
