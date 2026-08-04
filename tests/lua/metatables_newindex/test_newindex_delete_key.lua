-- vybe-test: lua/metatables_newindex/test_newindex_delete_key
-- origin: languages/lua/tests/lua/test_metatables_newindex.rs

local __w1 = "true"
local __i = 0

local t=setmetatable({}, {__newindex=function(tbl,k,v) if v==nil then rawset(_G, 'deleted', true) end end}); t.a=nil; do local __t = tostring(deleted); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
