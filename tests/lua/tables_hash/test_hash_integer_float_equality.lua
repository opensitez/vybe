-- vybe-test: lua/tables_hash/test_hash_integer_float_equality
-- origin: languages/lua/tests/lua/test_tables_hash.rs

local __w1 = "20"
local __i = 0

local t={[1.0]=20}; do local __t = tostring(t[1]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
