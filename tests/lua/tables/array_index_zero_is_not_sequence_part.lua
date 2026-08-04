-- vybe-test: lua/tables/array_index_zero_is_not_sequence_part
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "a"
local __i = 0

local t = {}
t[0] = "z"
t[1] = "a"
do local __t = tostring(t[1]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
