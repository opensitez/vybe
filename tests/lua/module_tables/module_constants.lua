-- vybe-test: lua/module_tables/module_constants
-- origin: languages/lua/tests/lua/test_module_tables.rs

local __w1 = "100,app"
local __i = 0

local Config = {max = 100, min = 0, name = "app"}
do local __t = tostring(Config.max - Config.min .. "," .. Config.name); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
