-- vybe-test: lua/oop/subclass_overrides_parent_method_field
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "c"
local __i = 0

local Base = {tag = "b"}
function Base:tag() return self.tag end
local child = setmetatable({tag = "c"}, {__index = Base})
do local __t = tostring(child:tag()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
