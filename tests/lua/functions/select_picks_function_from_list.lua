-- vybe-test: lua/functions/select_picks_function_from_list
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "2"
local __i = 0

local fns = {function() return 1 end, function() return 2 end}
do local __t = tostring(fns[2]()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
