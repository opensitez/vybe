-- vybe-test: lua/types_exhaustive/fn_closure_nested
-- origin: languages/lua/tests/lua/test_types_exhaustive.rs

local __w1 = "15"
local __i = 0

local function f(x)
  return function(y) return x + y end
end
do local __t = tostring(f(5)(10)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
