-- vybe-test: lua/closures_nested/test_nested_closure_deep_read
-- origin: languages/lua/tests/lua/test_closures_nested.rs

local __w1 = "1234"
local __i = 0

local function f1(a) return function(b) return function(c) return function(d) return a..b..c..d end end end end; do local __t = tostring(f1(1)(2)(3)(4)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
