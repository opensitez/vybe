-- vybe-test: lua/assignment/assignment_vararg_truncated_in_middle
-- origin: languages/lua/tests/lua/test_assignment.rs

local __w1 = "10,99"
local __i = 0

local function f(...)
  local a, b = ..., 99
  do local __t = tostring(a .. "," .. b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end
f(10, 20)

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
