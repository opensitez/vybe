-- vybe-test: lua/closures/closure_mutates_upvalue_readable_by_sibling
-- origin: languages/lua/tests/lua/test_closures.rs

local __w1 = "written"
local __i = 0

local val = 'initial'
local writer = function() val = 'written' end
local reader = function() return val end
writer()
do local __t = tostring(reader()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
