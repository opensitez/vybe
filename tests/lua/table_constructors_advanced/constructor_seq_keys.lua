-- vybe-test: lua/table_constructors_advanced/constructor_seq_keys
-- origin: languages/lua/tests/lua/test_table_constructors_advanced.rs

local __w1 = "ac"
local __i = 0

local t = {"a", "b", "c"}
do local __t = tostring(t[1] .. t[3]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
