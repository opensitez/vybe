-- vybe-test: lua/oop/instance_method_invokes_with_colon_syntax
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "4"
local __i = 0

local Acc = {}
function Acc.new() return setmetatable({n = 0}, {__index = Acc}) end
function Acc:add(x) self.n = self.n + x end
local a = Acc.new()
a:add(4)
do local __t = tostring(a.n); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
