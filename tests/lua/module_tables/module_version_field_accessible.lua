-- vybe-test: lua/module_tables/module_version_field_accessible
-- origin: languages/lua/tests/lua/test_module_tables.rs

local __w1 = "mymod@1.0.0"
local __i = 0

local M = {_VERSION = '1.0.0', name = 'mymod'}
do local __t = tostring(M.name .. '@' .. M._VERSION); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
