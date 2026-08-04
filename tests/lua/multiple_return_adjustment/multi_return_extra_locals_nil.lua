-- vybe-test: lua/multiple_return_adjustment/multi_return_extra_locals_nil
-- origin: languages/lua/tests/lua/test_multiple_return_adjustment.rs

local __w1 = "nil"
local __i = 0

local function f() return 1 end
local a, b = f()
do local __t = tostring(tostring(b)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
