-- vybe-test: lua/garbage_collection_advanced/test_gc_weak_tables_keys
-- origin: languages/lua/tests/lua/test_garbage_collection_advanced.rs

local __w1 = "0"
local __i = 0

local t = {}
setmetatable(t, {__mode = 'k'})
local k = {}
t[k] = 1
k = nil
collectgarbage('collect')
local count = 0
for k, v in pairs(t) do count = count + 1 end
do local __t = tostring(count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
