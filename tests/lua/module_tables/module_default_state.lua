-- vybe-test: lua/module_tables/module_default_state
-- origin: languages/lua/tests/lua/test_module_tables.rs

local __w1 = "30"
local __i = 0

local M = {values = {}}
function M.add(v) M.values[#M.values+1] = v end
function M.sum()
  local s = 0
  for _, v in ipairs(M.values) do s = s + v end
  return s
end
M.add(5); M.add(10); M.add(15)
do local __t = tostring(M.sum()); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
