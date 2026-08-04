-- vybe-test: lua/collectgarbage/collectgarbage_default_performs_full_collection
-- origin: languages/lua/tests/lua/test_collectgarbage.rs

local __w1 = "true"
local __i = 0

local ok = pcall(collectgarbage)
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
