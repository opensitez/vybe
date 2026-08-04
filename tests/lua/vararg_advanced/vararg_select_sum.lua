-- vybe-test: lua/vararg_advanced/vararg_select_sum
-- origin: languages/lua/tests/lua/test_vararg_advanced.rs

local __w1 = "15"
local __i = 0

local function sum(...)
  local s = 0
  for i = 1, select('#', ...) do
    s = s + select(i, ...)
  end
  return s
end
do local __t = tostring(sum(1, 2, 3, 4, 5)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
