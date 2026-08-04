-- vybe-test: lua/closures_upvalues/test_upvalue_closure_identity_same_scope
-- origin: languages/lua/tests/lua/test_closures_upvalues.rs

local __w1 = "false"
local __i = 0

local t={}; for i=1,2 do local a=1; local function f() return a end; t[i]=f end; do local __t = tostring(tostring(t[1]==t[2])); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
