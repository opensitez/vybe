-- vybe-test: lua/metatables_index/test_index_table_basic
-- origin: languages/lua/tests/lua/test_metatables_index.rs

local __w1 = "1"
local __i = 0

local t1={a=1}; local t2={}; setmetatable(t2, {__index=t1}); do local __t = tostring(t2.a); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
