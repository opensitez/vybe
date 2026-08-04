-- vybe-test: lua/table_sparse_behavior/test_table_sparse_behavior_negative
-- origin: languages/lua/tests/lua/test_table_sparse_behavior.rs

local __w1 = "true"
local __i = 0

local t = {[1] = 1, [100] = 8}
local c = 0
for _ in pairs(t) do c = c + 1 end
do local __t = tostring(c == 2); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
