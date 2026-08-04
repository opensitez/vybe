-- vybe-test: lua/module_tables/module_nested_calls
-- origin: languages/lua/tests/lua/test_module_tables.rs

local __w1 = "12"
local __i = 0

local M = {}
function M.double(n) return n * 2 end
function M.quadruple(n) return M.double(M.double(n)) end
do local __t = tostring(M.quadruple(3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
