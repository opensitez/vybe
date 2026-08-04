-- vybe-test: lua/metatables_newindex/test_newindex_table_basic
-- origin: languages/lua/tests/lua/test_metatables_newindex.rs

local __w1 = "1 nil"
local __i = 0

local t1={}; local t2={}; setmetatable(t2, {__newindex=t1}); t2.a=1; do local __t = tostring(t1.a..' '..(t2.a or 'nil')); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
