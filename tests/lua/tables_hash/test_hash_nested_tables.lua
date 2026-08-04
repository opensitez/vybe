-- vybe-test: lua/tables_hash/test_hash_nested_tables
-- origin: languages/lua/tests/lua/test_tables_hash.rs

local __w1 = "1"
local __i = 0

local t={a={b={c=1}}}; do local __t = tostring(t.a.b.c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
