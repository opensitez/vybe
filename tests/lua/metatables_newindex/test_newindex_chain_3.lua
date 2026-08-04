-- vybe-test: lua/metatables_newindex/test_newindex_chain_3
-- origin: languages/lua/tests/lua/test_metatables_newindex.rs

local __w1 = "99"
local __i = 0

local t1={}; local t2=setmetatable({}, {__newindex=t1}); local t3=setmetatable({}, {__newindex=t2}); local t4=setmetatable({}, {__newindex=t3}); t4.x=99; do local __t = tostring(t1.x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
