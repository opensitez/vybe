-- vybe-test: lua/metatables/__add_with_two_operands_of_same_metatable
-- origin: languages/lua/tests/lua/test_metatables.rs

local __w1 = "5"
local __i = 0

local mt = {__add = function(a, b) return a.n + b.n end}
local x = setmetatable({n = 2}, mt)
local y = setmetatable({n = 3}, mt)
do local __t = tostring(x + y); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
