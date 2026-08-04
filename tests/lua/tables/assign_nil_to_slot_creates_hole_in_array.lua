-- vybe-test: lua/tables/assign_nil_to_slot_creates_hole_in_array
-- origin: languages/lua/tests/lua/test_tables.rs

local __w1 = "true"
local __i = 0

local t = {1, 2, 3}
t[2] = nil
do local __t = tostring(t[2] == nil); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
