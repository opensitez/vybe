-- vybe-test: lua/metamethods/metamethod_le_orders_values
-- origin: languages/lua/tests/lua/test_metamethods.rs

local __w1 = "true"
local __i = 0

local mt={__le=function(a,b) return a.n<=b.n end}
do local __t = tostring(setmetatable({n=2},mt)<=setmetatable({n=3},mt)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
