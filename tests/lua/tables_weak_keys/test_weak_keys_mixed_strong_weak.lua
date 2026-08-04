-- vybe-test: lua/tables_weak_keys/test_weak_keys_mixed_strong_weak
-- origin: languages/lua/tests/lua/test_tables_weak_keys.rs

local __w1 = "nil 2"
local __i = 0

local t=setmetatable({}, {__mode='k'}); local k1={}; local k2={}; t[k1]=1; t[k2]=2; k1=nil; collectgarbage(); do local __t = tostring((t[k1] or 'nil')..' '..t[k2]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
