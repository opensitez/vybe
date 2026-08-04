-- vybe-test: lua/functions/function_passed_table_and_reads_field
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "9"
local __i = 0

function read_x(t) return t.x end
do local __t = tostring(read_x({x = 9})); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
