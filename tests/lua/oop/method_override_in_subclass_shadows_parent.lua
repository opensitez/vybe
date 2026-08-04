-- vybe-test: lua/oop/method_override_in_subclass_shadows_parent
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "child"
local __i = 0

local Base = {}
Base.__index = Base
function Base:describe() return 'base' end
local Child = setmetatable({}, {__index = Base})
Child.__index = Child
function Child:describe() return 'child' end
local obj = setmetatable({}, Child)
do local __t = tostring(obj:describe()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
