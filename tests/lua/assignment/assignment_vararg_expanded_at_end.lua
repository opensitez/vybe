-- vybe-test: lua/assignment/assignment_vararg_expanded_at_end
-- origin: languages/lua/tests/lua/test_assignment.rs

local __w1 = "5,10,20"
local __i = 0

local function f(...)
  local a, b, c = 5, ...
  do local __t = tostring(a .. "," .. b .. "," .. c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end
f(10, 20)

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
