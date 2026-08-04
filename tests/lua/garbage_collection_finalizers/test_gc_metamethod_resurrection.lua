-- vybe-test: lua/garbage_collection_finalizers/test_gc_metamethod_resurrection
-- origin: languages/lua/tests/lua/test_garbage_collection_finalizers.rs

local __w1 = "42"
local __i = 0

local res; local t = setmetatable({a=42}, {__gc = function(obj) res = obj end}); t = nil; collectgarbage('collect'); do local __t = tostring(res.a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
