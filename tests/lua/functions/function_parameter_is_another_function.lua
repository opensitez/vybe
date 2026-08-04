-- vybe-test: lua/functions/function_parameter_is_another_function
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "8"
local __i = 0

local function twice_call(f) return f() + f() end
do local __t = tostring(twice_call(function() return 4 end)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
