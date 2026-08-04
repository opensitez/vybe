-- vybe-test: lua/functions/extra_arguments_are_ignored
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "1"
local __i = 0

function one(x) return x end
do local __t = tostring(one(1, 2, 3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
