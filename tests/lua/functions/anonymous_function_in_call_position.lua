-- vybe-test: lua/functions/anonymous_function_in_call_position
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "7"
local __i = 0

local function apply(f, x) return f(x) end
do local __t = tostring(apply(function(v) return v + 5 end, 2)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
