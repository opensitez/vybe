-- vybe-test: lua/oop/mixin_copies_methods_into_target_class
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "eagle flies"
local __i = 0

local function mixin(target, source)
  for k, v in pairs(source) do target[k] = v end
end
local Fly = {fly = function(self) return self.name .. ' flies' end}
local Bird = {name = 'eagle'}
mixin(Bird, Fly)
do local __t = tostring(Bird:fly()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
