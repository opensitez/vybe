-- vybe-test: lua/closures_upvalues/test_upvalue_escaping_scope_goto
-- origin: languages/lua/tests/lua/test_closures_upvalues.rs

local __w1 = "2"
local __i = 0

local f; do local a=1; ::lbl::; f=function() return a end; if a==1 then a=2; goto lbl end end; do local __t = tostring(f()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
