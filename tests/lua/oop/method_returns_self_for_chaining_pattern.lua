-- vybe-test: lua/oop/method_returns_self_for_chaining_pattern
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "ab"
local __i = 0

local Builder = {}
function Builder.new() return setmetatable({}, {__index = Builder}) end
function Builder:add(x)
  self.parts = (self.parts or "") .. x
  return self
end
local b = Builder.new():add("a"):add("b")
do local __t = tostring(b.parts); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
