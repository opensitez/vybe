-- vybe-test: lua/tables_weak_values/test_weak_values_strong_keys
-- origin: languages/lua/tests/lua/test_tables_weak_values.rs

local __w1 = "0"
local __i = 0

local t=setmetatable({}, {__mode='v'}); local k={}; local v={}; t[k]=v; v=nil; collectgarbage(); local cnt=0; for _ in pairs(t) do cnt=cnt+1 end; do local __t = tostring(cnt); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
