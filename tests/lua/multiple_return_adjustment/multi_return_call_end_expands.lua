-- vybe-test: lua/multiple_return_adjustment/multi_return_call_end_expands
-- origin: languages/lua/tests/lua/test_multiple_return_adjustment.rs

local __w1 = "6"
local __i = 0

local function f() return 2, 3 end
local function add(a, b, c) return a + b + c end
do local __t = tostring(add(1, f())); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
