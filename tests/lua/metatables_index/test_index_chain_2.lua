-- vybe-test: lua/metatables_index/test_index_chain_2
-- origin: languages/lua/tests/lua/test_metatables_index.rs

local __w1 = "10"
local __i = 0

local t1={a=10}; local t2={}; setmetatable(t2, {__index=t1}); local t3={}; setmetatable(t3, {__index=t2}); do local __t = tostring(t3.a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
