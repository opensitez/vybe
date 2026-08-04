-- vybe-test: lua/oop/tostring_metamethod_on_instance
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "(3,4)"
local __i = 0

local Vec = {}
function Vec.new(x, y) return setmetatable({x=x, y=y}, {__tostring = function(v) return '(' .. v.x .. ',' .. v.y .. ')' end}) end
local v = Vec.new(3, 4)
do local __t = tostring(tostring(v)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
