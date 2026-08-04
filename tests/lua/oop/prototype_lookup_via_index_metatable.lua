-- vybe-test: lua/oop/prototype_lookup_via_index_metatable
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "hi"
local __i = 0

local proto = {greet = function() return "hi" end}
local obj = setmetatable({}, {__index = proto})
do local __t = tostring(obj:greet()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
