-- vybe-test: lua/functions/missing_arguments_become_nil
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "nil"
local __i = 0

function show(x) do local __t = tostring(tostring(x)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end end
show()

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
