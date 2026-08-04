-- vybe-test: lua/tables_weak_keys/test_weak_keys_ephemeron_behavior
-- origin: languages/lua/tests/lua/test_tables_weak_keys.rs

local __w1 = "0"
local __i = 0

local t=setmetatable({}, {__mode='k'}); local k={}; t[k]=k; k=nil; collectgarbage(); local cnt=0; for _ in pairs(t) do cnt=cnt+1 end; do local __t = tostring(cnt); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
