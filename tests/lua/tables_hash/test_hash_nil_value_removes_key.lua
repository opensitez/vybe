-- vybe-test: lua/tables_hash/test_hash_nil_value_removes_key
-- origin: languages/lua/tests/lua/test_tables_hash.rs

local __w1 = "0"
local __i = 0

local t={a=1}; t.a=nil; local cnt=0; for k,v in pairs(t) do cnt=cnt+1 end; do local __t = tostring(cnt); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
