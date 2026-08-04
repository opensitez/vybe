-- vybe-test: lua/closures_ext/test_closure_upvalue_table
-- origin: languages/lua/tests/lua/test_closures_ext.rs

local __w1 = "2"
local __i = 0

local t={a=1}; local function f() t.a=t.a+1 return t.a end; f(); do local __t = tostring(t.a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
