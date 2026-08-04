-- vybe-test: lua/vararg_functions_advanced/vararg_sum_recursive
-- origin: languages/lua/tests/lua/test_vararg_functions_advanced.rs

local __w1 = "15"
local __i = 0

local function sum(head, ...)
  if not head then return 0 end
  return head + sum(...)
end
do local __t = tostring(sum(1, 2, 3, 4, 5)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
