-- vybe-test: lua/tables_hash/test_hash_function_keys
-- origin: languages/lua/tests/lua/test_tables_hash.rs

local __w1 = "3"
local __i = 0

local f1=function() end; local f2=function() end; local t={[f1]=1, [f2]=2}; do local __t = tostring(t[f1]+t[f2]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
