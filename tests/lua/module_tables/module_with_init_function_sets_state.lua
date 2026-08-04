-- vybe-test: lua/module_tables/module_with_init_function_sets_state
-- origin: languages/lua/tests/lua/test_module_tables.rs

local __w1 = "true,42"
local __i = 0

local M = {ready = false}
function M.init() M.ready = true; M.value = 42 end
M.init()
do local __t = tostring(tostring(M.ready) .. ',' .. M.value); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
