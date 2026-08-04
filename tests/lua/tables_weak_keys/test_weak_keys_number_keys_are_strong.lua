-- vybe-test: lua/tables_weak_keys/test_weak_keys_number_keys_are_strong
-- origin: languages/lua/tests/lua/test_tables_weak_keys.rs

local __w1 = "1"
local __i = 0

local t=setmetatable({}, {__mode='k'}); local k=42; t[k]=1; k=nil; collectgarbage(); do local __t = tostring(t[42]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
