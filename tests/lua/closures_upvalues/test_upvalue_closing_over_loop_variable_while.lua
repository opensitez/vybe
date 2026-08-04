-- vybe-test: lua/closures_upvalues/test_upvalue_closing_over_loop_variable_while
-- origin: languages/lua/tests/lua/test_closures_upvalues.rs

local __w1 = "123"
local __i = 0

local t={}; local i=1; while i<=3 do local j=i; t[i]=function() return j end; i=i+1 end; do local __t = tostring(t[1]()..t[2]()..t[3]()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
