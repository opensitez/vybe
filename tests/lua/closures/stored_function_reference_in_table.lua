-- vybe-test: lua/closures/stored_function_reference_in_table
-- origin: languages/lua/tests/lua/test_closures.rs

local __w1 = "9"
local __i = 0

local api = { run = function() return 9 end }
do local __t = tostring(api.run()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
