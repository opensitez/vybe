-- vybe-test: lua/table_unpack_range/unpack_to_args
-- origin: languages/lua/tests/lua/test_table_unpack_range.rs

local __w1 = "6"
local __i = 0

local function add(a, b, c) return a + b + c end
local t = {1, 2, 3}
do local __t = tostring(add(table.unpack(t))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
