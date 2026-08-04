-- vybe-test: lua/metatables_index/test_index_function_receives_table
-- origin: languages/lua/tests/lua/test_metatables_index.rs

local __w1 = "1 true"
local __i = 0

local target; local t={}; setmetatable(t, {__index=function(tbl, k) target=tbl; return 1 end}); do local __t = tostring(t.a..' '..tostring(t==target)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
