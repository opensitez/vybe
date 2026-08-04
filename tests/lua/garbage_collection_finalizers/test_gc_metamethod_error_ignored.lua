-- vybe-test: lua/garbage_collection_finalizers/test_gc_metamethod_error_ignored
-- origin: languages/lua/tests/lua/test_garbage_collection_finalizers.rs

local __w1 = "true"
local __i = 0

local t = setmetatable({}, {__gc = function() error('boom') end}); t = nil; local ok = pcall(function() collectgarbage('collect') end); do local __t = tostring(tostring(ok)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
