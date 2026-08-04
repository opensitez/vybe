-- vybe-test: lua/programs/module_style_return_table
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "5"
local __i = 0

local M = {}
function M.add(a, b) return a + b end
do local __t = tostring(M.add(2, 3)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
