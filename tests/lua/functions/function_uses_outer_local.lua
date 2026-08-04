-- vybe-test: lua/functions/function_uses_outer_local
-- origin: languages/lua/tests/lua/test_functions.rs

local __w1 = "105"
local __i = 0

local base = 100
function add(x) return base + x end
do local __t = tostring(add(5)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
