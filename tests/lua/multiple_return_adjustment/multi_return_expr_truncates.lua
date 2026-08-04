-- vybe-test: lua/multiple_return_adjustment/multi_return_expr_truncates
-- origin: languages/lua/tests/lua/test_multiple_return_adjustment.rs

local __w1 = "1"
local __i = 0

local function f() return 1, 2, 3 end
local x = f() + 0
do local __t = tostring(x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
