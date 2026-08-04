-- vybe-test: lua/metatables_index/test_index_metamethod_type_nil
-- origin: languages/lua/tests/lua/test_metatables_index.rs

local __w1 = "nil_foo"
local __i = 0

debug.setmetatable(nil, {__index=function(n, k) return 'nil_'..k end}); local x=nil; do local __t = tostring(x['foo']); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
