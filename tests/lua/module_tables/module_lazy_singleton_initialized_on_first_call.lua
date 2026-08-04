-- vybe-test: lua/module_tables/module_lazy_singleton_initialized_on_first_call
-- origin: languages/lua/tests/lua/test_module_tables.rs

local __w1 = "3"
local __i = 0

local instance = nil
local M = {}
function M.get()
  if not instance then instance = {count = 0} end
  instance.count = instance.count + 1
  return instance
end
M.get(); M.get()
do local __t = tostring(M.get().count); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
