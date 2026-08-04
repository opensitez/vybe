-- vybe-test: lua/closures_upvalues/test_upvalue_closure_identity_same_outer_function
-- origin: languages/lua/tests/lua/test_closures_upvalues.rs

local __w1 = "true"
local __i = 0

local function outer() local function inner() end; return inner, inner end; local f1, f2 = outer(); do local __t = tostring(tostring(f1==f2)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
