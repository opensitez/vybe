-- vybe-test: lua/assignment/multiple_assignment_with_empty_vararg
-- origin: languages/lua/tests/lua/test_assignment.rs

local __w1 = "1,nil"
local __i = 0

local function f(...)
  local a, b = 1, ...
  do local __t = tostring(tostring(a) .. "," .. tostring(b)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end
end
f()

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
