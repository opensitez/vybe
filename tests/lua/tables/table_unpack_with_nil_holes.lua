-- vybe-test: lua/tables/table_unpack_with_nil_holes
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "10,nil,30"
local __i = 0

local a, b, c = table.unpack({10, nil, 30}, 1, 3)
do local __t = tostring(tostring(a) .. "," .. tostring(b) .. "," .. tostring(c)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
