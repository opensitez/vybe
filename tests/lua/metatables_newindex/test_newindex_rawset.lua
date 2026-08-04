-- vybe-test: lua/metatables_newindex/test_newindex_rawset
-- origin: languages/lua/tests/lua/test_metatables_newindex.rs

local __w1 = "1"
local __i = 0

local t=setmetatable({}, {__newindex=function() error('boom') end}); rawset(t, 'a', 1); do local __t = tostring(t.a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
