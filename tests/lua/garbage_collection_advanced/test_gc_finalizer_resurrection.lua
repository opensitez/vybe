-- vybe-test: lua/garbage_collection_advanced/test_gc_finalizer_resurrection
-- origin: languages/lua/tests/lua/test_garbage_collection_advanced.rs

local __w1 = "table"
local __i = 0

local resurrected
local t = {}
setmetatable(t, {__gc = function(o) resurrected = o end})
t = nil
collectgarbage('collect')
do local __t = tostring(type(resurrected)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
