-- vybe-test: lua/collectgarbage/collectgarbage_restart_after_stop
-- origin: languages/lua/tests/lua/test_collectgarbage.rs

local __w1 = "true"
local __i = 0

collectgarbage("stop")
local was = collectgarbage("isrunning")
collectgarbage("restart")
do local __t = tostring(was == false and collectgarbage("isrunning")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
