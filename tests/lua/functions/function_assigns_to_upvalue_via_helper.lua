-- vybe-test: lua/functions/function_assigns_to_upvalue_via_helper
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "7"
local __i = 0

local total = 0
function add_to_total(x) total = total + x end
add_to_total(3)
add_to_total(4)
do local __t = tostring(total); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
