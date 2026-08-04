-- vybe-test: lua/collectgarbage/collectgarbage_setpause_modifies_pause_and_returns_previous
-- origin: languages/lua/tests/lua/test_collectgarbage.rs

local __w1 = "true"
local __i = 0

local prev = collectgarbage("setpause", 150)
local cur = collectgarbage("setpause", prev)
do local __t = tostring(type(prev) == "number" and cur == 150); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
