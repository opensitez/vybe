-- vybe-test: lua/tables/hash_part_does_not_affect_length_operator
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "2"
local __i = 0

local t = {1, 2}
t.hidden = 99
do local __t = tostring(#t); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
