-- vybe-test: lua/tables_hash/test_hash_self_reference
-- origin: languages/lua/tests/lua/test_tables_hash.rs

local __w1 = "true"
local __i = 0

local t={}; t.self=t; do local __t = tostring(t.self.self == t); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
