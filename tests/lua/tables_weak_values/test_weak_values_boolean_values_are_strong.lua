-- vybe-test: lua/tables_weak_values/test_weak_values_boolean_values_are_strong
-- origin: languages/lua/tests/lua/test_tables_weak_values.rs

local __w1 = "true"
local __i = 0

local t=setmetatable({}, {__mode='v'}); local v=true; t[1]=v; v=nil; collectgarbage(); do local __t = tostring(t[1]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
