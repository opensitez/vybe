-- vybe-test: lua/tables_weak_values/test_weak_values_mixed_strong_weak
-- origin: languages/lua/tests/lua/test_tables_weak_values.rs

local __w1 = "nil true"
local __i = 0

local t=setmetatable({}, {__mode='v'}); local v1={}; local v2={}; t[1]=v1; t[2]=v2; v1=nil; collectgarbage(); do local __t = tostring((t[1] or 'nil')..' '..tostring(t[2]==v2)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
