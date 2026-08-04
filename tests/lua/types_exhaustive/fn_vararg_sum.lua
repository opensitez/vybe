-- vybe-test: lua/types_exhaustive/fn_vararg_sum
-- origin: languages/lua/tests/lua/test_types_exhaustive.rs

local __w1 = "3"
local __i = 0

local function f(...) return select("#", ...) end
do local __t = tostring(f(1, 2, 3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
