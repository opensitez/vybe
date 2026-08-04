-- vybe-test: lua/assignment/assign_from_function_multiple_returns_to_locals
-- origin: languages/lua/tests/lua/test_assignment.rs

local __w1 = "1,3"
local __i = 0

local function minmax(a, b)
  if a < b then return a, b else return b, a end
end
local lo, hi = minmax(3, 1)
do local __t = tostring(lo .. "," .. hi); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
