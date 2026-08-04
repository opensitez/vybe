-- vybe-test: lua/basics/function_returns_to_print_directly
-- origin: languages/lua/tests/lua/test_basics.rs

local __w1 = "2"
local __i = 0

function two() return 2 end
do local __t = tostring(two()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
