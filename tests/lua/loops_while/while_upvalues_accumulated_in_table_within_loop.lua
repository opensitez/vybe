-- vybe-test: lua/loops_while/while_upvalues_accumulated_in_table_within_loop
-- origin: languages/lua/tests/lua/test_loops_while.rs

local __w1 = "10,20,30,"
local __i = 0

local out = ''
local i = 1
while i <= 3 do
  local x = i * 10
  out = out .. x .. ','
  i = i + 1
end
do local __t = tostring(out); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
