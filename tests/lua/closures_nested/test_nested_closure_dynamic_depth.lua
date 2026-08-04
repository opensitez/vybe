-- vybe-test: lua/closures_nested/test_nested_closure_dynamic_depth
-- origin: languages/lua/tests/lua/test_closures_nested.rs

local __w1 = "10"
local __i = 0

local function make_adder(n) local sum=n; local function adder(x) if x then sum=sum+x; return adder else return sum end end; return adder end; do local __t = tostring(make_adder(1)(2)(3)(4)()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
