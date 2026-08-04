-- vybe-test: lua/iteration/next_on_empty_table_returns_nil
-- origin: languages/lua/tests/lua/test_iteration.rs

local __w1 = "nil"
local __i = 0

local k = next({})
do local __t = tostring(tostring(k)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
