-- vybe-test: lua/garbage_collection_advanced/test_gc_step_execution
-- origin: languages/lua/tests/lua/test_garbage_collection_advanced.rs

local __w1 = "true"
local __i = 0

local pre = collectgarbage('count')
local t = {}
for i = 1, 10000 do t[i] = tostring(i) end
t = nil
local b = collectgarbage('step')
local post = collectgarbage('count')
do local __t = tostring(type(b) == 'boolean'); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
