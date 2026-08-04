-- vybe-test: lua/oop/method_table_on_prototype_not_copied_to_instance
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "1"
local __i = 0

local Proto = {val = 1}
local a = setmetatable({}, {__index = Proto})
local b = setmetatable({}, {__index = Proto})
a.val = 2
do local __t = tostring(b.val); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
