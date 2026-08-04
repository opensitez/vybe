-- vybe-test: lua/globals/deleting_global_by_assigning_nil
-- origin: languages/lua/tests/lua/test_globals.rs

local __w1 = "nil"
local __i = 0

some_global = 'hello'
some_global = nil
do local __t = tostring(tostring(some_global)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
