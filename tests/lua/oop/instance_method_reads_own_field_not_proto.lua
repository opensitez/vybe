-- vybe-test: lua/oop/instance_method_reads_own_field_not_proto
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "inst"
local __i = 0

local proto = {name = "base"}
local obj = setmetatable({name = "inst"}, {__index = proto})
function obj:label() return self.name end
do local __t = tostring(obj:label()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
