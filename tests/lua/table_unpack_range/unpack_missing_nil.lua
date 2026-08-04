-- vybe-test: lua/table_unpack_range/unpack_missing_nil
-- origin: languages/lua/tests/lua/test_table_unpack_range.rs

local __w1 = "nil"
local __i = 0

local a, b, c = table.unpack({1, 2})
do local __t = tostring(tostring(c)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
