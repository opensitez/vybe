-- vybe-test: lua/garbage_collection_finalizers/test_gc_metamethod_called
-- origin: languages/lua/tests/lua/test_garbage_collection_finalizers.rs

local __w1 = "true"
local __i = 0

local finalized = false; local t = setmetatable({}, {__gc = function() finalized = true end}); t = nil; collectgarbage('collect'); do local __t = tostring(tostring(finalized)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
