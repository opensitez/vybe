-- vybe-test: lua/closures/closure_over_loop_variable_requires_iife_pattern
-- origin: languages/lua/tests/lua/test_closures.rs

local __w1 = "3"
local __i = 0

local fns={}
for i=1,2 do
  fns[i]=function() return i end
end
do local __t = tostring(fns[1]()+fns[2]()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
