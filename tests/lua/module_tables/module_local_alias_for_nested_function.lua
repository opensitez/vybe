-- vybe-test: lua/module_tables/module_local_alias_for_nested_function
-- origin: languages/lua/tests/lua/test_module_tables.rs

local __w1 = "3"
local __i = 0

local M = {}
local math_floor = math.floor
function M.truncate(x) return math_floor(x) end
do local __t = tostring(M.truncate(3.9)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
