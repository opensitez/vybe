-- vybe-test: lua/functions/anonymous_function_captured_in_loop
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "30"
local __i = 0

local t = {}
for i = 1, 2 do t[i] = function() return i * 10 end end
do local __t = tostring(t[1]() + t[2]()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
