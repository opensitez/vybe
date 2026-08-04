-- vybe-test: lua/table_constructors/constructor_duplicate_keys_last_wins
-- origin: languages/lua/tests/lua/test_table_constructors.rs

local __w1 = "2"
local __i = 0

local t = {a=1, a=2}
do local __t = tostring(t.a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
