-- vybe-test: lua/metatables_index/test_index_chain_function
-- origin: languages/lua/tests/lua/test_metatables_index.rs

local __w1 = "z1"
local __i = 0

local t1=setmetatable({}, {__index=function(t,k) return k..'1' end}); local t2=setmetatable({}, {__index=t1}); do local __t = tostring(t2.z); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
