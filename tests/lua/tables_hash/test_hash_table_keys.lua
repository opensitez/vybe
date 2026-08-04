-- vybe-test: lua/tables_hash/test_hash_table_keys
-- origin: languages/lua/tests/lua/test_tables_hash.rs

local __w1 = "v1 v2"
local __i = 0

local k1={}; local k2={}; local t={[k1]='v1', [k2]='v2'}; do local __t = tostring(t[k1]..' '..t[k2]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
