-- vybe-test: lua/garbage_collection_api/test_gc_stop_restart
-- origin: languages/lua/tests/lua/test_garbage_collection_api.rs

local __w1 = "ok"
local __i = 0

collectgarbage('stop'); collectgarbage('restart'); do local __t = tostring('ok'); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
