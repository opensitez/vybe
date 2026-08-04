-- vybe-test: lua/collectgarbage/collectgarbage_isrunning_returns_boolean
-- origin: languages/lua/tests/lua/test_collectgarbage.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(type(collectgarbage("isrunning")) == "boolean"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
