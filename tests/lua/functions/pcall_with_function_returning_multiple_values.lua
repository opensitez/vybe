-- vybe-test: lua/functions/pcall_with_function_returning_multiple_values
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "true,10,20,30"
local __i = 0

local ok, a, b, c = pcall(function() return 10, 20, 30 end)
do local __t = tostring(tostring(ok) .. ',' .. a .. ',' .. b .. ',' .. c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
