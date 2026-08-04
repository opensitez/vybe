-- vybe-test: lua/loops_repeat_guard/test_loops_repeat_guard_product_until_limit
-- origin: languages/lua/tests/lua/test_loops_repeat_guard.rs

local __w1 = "32"
local __i = 0

local n = 1
repeat n = n * 2; if n > 20 then break end until false
do local __t = tostring(n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
