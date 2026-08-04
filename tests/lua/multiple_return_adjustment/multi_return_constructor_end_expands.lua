-- vybe-test: lua/multiple_return_adjustment/multi_return_constructor_end_expands
-- origin: languages/lua/tests/lua/test_multiple_return_adjustment.rs

local __w1 = "4"
local __i = 0

local function f() return 10, 20 end
local t = {1, 2, f()}
do local __t = tostring(#t); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
