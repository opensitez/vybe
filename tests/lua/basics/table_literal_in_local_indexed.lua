-- vybe-test: lua/basics/table_literal_in_local_indexed
-- origin: languages/lua/tests/lua/test_basics.rs

local __w1 = "ab"
local __i = 0

local t = {"a", "b"}
do local __t = tostring(t[1] .. t[2]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
