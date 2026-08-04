-- vybe-test: lua/tables_hash/test_hash_number_keys
-- origin: languages/lua/tests/lua/test_tables_hash.rs

local __w1 = "30"
local __i = 0

local t={[1.5]=10, [2.5]=20}; do local __t = tostring(t[1.5]+t[2.5]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
