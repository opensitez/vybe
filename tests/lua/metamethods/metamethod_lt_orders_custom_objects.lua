-- vybe-test: lua/metamethods/metamethod_lt_orders_custom_objects
-- origin: languages/lua/tests/lua/test_metamethods.rs

local __w1 = "true"
local __i = 0

local mt = {__lt = function(a, b) return a.score < b.score end}
do local __t = tostring(setmetatable({score=1}, mt) < setmetatable({score=2}, mt)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
