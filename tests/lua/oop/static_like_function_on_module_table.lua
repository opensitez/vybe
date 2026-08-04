-- vybe-test: lua/oop/static_like_function_on_module_table
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "8"
local __i = 0

local Math2 = {}
function Math2.twice(x) return x * 2 end
do local __t = tostring(Math2.twice(4)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
