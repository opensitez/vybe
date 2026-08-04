-- vybe-test: lua/oop/inheritance_child_falls_back_to_parent_method
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "p"
local __i = 0

local Parent = {kind = "p"}
function Parent:kind() return self.kind end
local child = setmetatable({}, {__index = Parent})
do local __t = tostring(child:kind()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
