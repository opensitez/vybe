-- vybe-test: lua/module_tables/module_extends_another_by_copying_methods
-- origin: languages/lua/tests/lua/test_module_tables.rs

local __w1 = "hi,bye"
local __i = 0

local Base = {greet = function() return 'hi' end}
local Child = {}
for k, v in pairs(Base) do Child[k] = v end
function Child.farewell() return 'bye' end
do local __t = tostring(Child.greet() .. ',' .. Child.farewell()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
