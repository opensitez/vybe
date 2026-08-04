-- vybe-test: lua/tables_mixed_ext/test_mixed_table_pairs
-- origin: languages/lua/tests/lua/test_tables_mixed_ext.rs

local __w1 = "1"
local __i = 0

local t={a=1}; local c=0; for k,v in pairs(t) do c=c+1 end; do local __t = tostring(c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
