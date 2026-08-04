-- vybe-test: lua/closures/two_closures_share_same_upvalue_cell
-- origin: languages/lua/tests/lua/test_closures.rs

local __w1 = "3"
local __i = 0

local x = 0
local inc = function() x = x + 1 end
local get = function() return x end
inc(); inc(); inc()
do local __t = tostring(get()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
