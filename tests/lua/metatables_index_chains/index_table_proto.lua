-- vybe-test: lua/metatables_index_chains/index_table_proto
-- origin: languages/lua/tests/lua/test_metatables_index_chains.rs

local __w1 = "hi world"
local __i = 0

local proto = {kind = "base", greet = function(self) return "hi " .. self.name end}
local obj = setmetatable({name = "world"}, {__index = proto})
do local __t = tostring(obj:greet()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
