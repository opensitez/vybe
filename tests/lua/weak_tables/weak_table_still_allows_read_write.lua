-- vybe-test: lua/weak_tables/weak_table_still_allows_read_write
-- origin: languages/lua/tests/lua/test_weak_tables.rs

local __w1 = "1"
local __i = 0

local t = setmetatable({}, {__mode = "k"})
t.x = 1
do local __t = tostring(t.x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
