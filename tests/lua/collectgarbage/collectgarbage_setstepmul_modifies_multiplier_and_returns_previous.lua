-- vybe-test: lua/collectgarbage/collectgarbage_setstepmul_modifies_multiplier_and_returns_previous
-- origin: languages/lua/tests/lua/test_collectgarbage.rs

local __w1 = "true"
local __i = 0

local prev = collectgarbage("setstepmul", 250)
local cur = collectgarbage("setstepmul", prev)
do local __t = tostring(type(prev) == "number" and cur == 250); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
