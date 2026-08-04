-- vybe-test: lua/assignment/update_table_field_with_self_reference
-- origin: languages/lua/tests/lua/test_assignment.rs

local __w1 = "2"
local __i = 0

local cfg = {count = 1}
cfg.count = cfg.count + 1
do local __t = tostring(cfg.count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
