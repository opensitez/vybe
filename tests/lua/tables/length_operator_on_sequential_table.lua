-- vybe-test: lua/tables/length_operator_on_sequential_table
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "3"
local __i = 0

local t = {}
t[1] = 1
t[2] = 2
t[3] = 3
do local __t = tostring(#t); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
