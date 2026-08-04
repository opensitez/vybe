-- vybe-test: lua/metatables_extended/meta_index_table
-- origin: languages/lua/tests/lua/test_metatables_extended.rs

local __w1 = "10"
local __i = 0

local parent = {x=10}
local child = setmetatable({}, {__index = parent})
do local __t = tostring(child.x); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
