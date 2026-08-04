-- vybe-test: lua/tables_hash/test_hash_large_number_of_keys
-- origin: languages/lua/tests/lua/test_tables_hash.rs

local __w1 = "50"
local __i = 0

local t={}; for i=1,100 do t['k'..i]=i end; do local __t = tostring(t.k50); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
