-- vybe-test: lua/closures/each_closure_has_distinct_upvalue_binding
-- origin: languages/lua/tests/lua/test_closures.rs

local __w1 = "1"
local __i = 0

local function make()
  local n=0
  return function() n=n+1 return n end
end
local a=make()
local b=make()
a()
do local __t = tostring(b()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
