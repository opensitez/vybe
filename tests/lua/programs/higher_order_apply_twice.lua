-- vybe-test: lua/programs/higher_order_apply_twice
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "5"
local __i = 0

local function apply_twice(f, x) return f(f(x)) end
do local __t = tostring(apply_twice(function(n) return n + 1 end, 3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
