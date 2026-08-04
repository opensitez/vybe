-- vybe-test: lua/collectgarbage/collectgarbage_garbage_creation_increases_memory_count
-- origin: languages/lua/tests/lua/test_collectgarbage.rs

local __w1 = "true"
local __i = 0

local before = collectgarbage("count")
local t = {}
for i=1,1000 do t[i] = {x=i} end
local after = collectgarbage("count")
do local __t = tostring(after > before); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
