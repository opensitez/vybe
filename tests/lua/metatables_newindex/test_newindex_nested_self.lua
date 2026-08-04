-- vybe-test: lua/metatables_newindex/test_newindex_nested_self
-- origin: languages/lua/tests/lua/test_metatables_newindex.rs

local __w1 = "1"
local __i = 0

local t={}; t.self=t; setmetatable(t, {__newindex=function(tbl, k, v) rawset(tbl, k, v) end}); t.a=1; do local __t = tostring(t.self.a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
