-- vybe-test: lua/module_tables/module_greet
-- origin: languages/lua/tests/lua/test_module_tables.rs

local __w1 = "hello lua"
local __i = 0

local M = {}
function M.greet(name) return "hello " .. name end
do local __t = tostring(M.greet("lua")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
