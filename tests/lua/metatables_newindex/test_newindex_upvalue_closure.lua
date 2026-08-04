-- vybe-test: lua/metatables_newindex/test_newindex_upvalue_closure
-- origin: languages/lua/tests/lua/test_metatables_newindex.rs

local __w1 = "2"
local __i = 0

local count=0; local t=setmetatable({}, {__newindex=function(tbl,k,v) count=count+1; rawset(tbl,k,v) end}); t.a=1; t.b=2; t.a=3; do local __t = tostring(count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
