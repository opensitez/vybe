-- vybe-test: lua/metatables_newindex/test_newindex_chain_function
-- origin: languages/lua/tests/lua/test_metatables_newindex.rs

local __w1 = "z 42"
local __i = 0

local k1, v1; local t1=setmetatable({}, {__newindex=function(t,k,v) k1=k; v1=v end}); local t2=setmetatable({}, {__newindex=t1}); t2.z=42; do local __t = tostring(k1..' '..v1); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
