-- vybe-test: lua/table_library_extended/table_unpack_defaults
-- origin: languages/lua/tests/lua/test_table_library_extended.rs

local __w1 = "10,20"
local __i = 0

local a, b = table.unpack({10, 20})
do local __t = tostring(a .. "," .. b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
