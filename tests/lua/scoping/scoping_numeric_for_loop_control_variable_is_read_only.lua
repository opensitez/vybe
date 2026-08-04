-- vybe-test: lua/scoping/scoping_numeric_for_loop_control_variable_is_read_only
-- origin: languages/lua/tests/lua/test_scoping.rs

local __w1 = "6"
local __i = 0

local sum = 0
for i = 1, 3 do
  sum = sum + i
  i = 100
end
do local __t = tostring(sum); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
