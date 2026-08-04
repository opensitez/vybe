-- vybe-test: lua/loops_for_numeric/numeric_for_control_var_read_from_closure_but_not_mutated
-- origin: languages/lua/tests/lua/test_loops_for_numeric.rs

local __w1 = "5"
local __i = 0

local last = 0
for i = 1, 5 do
  local function capture() return i end
  last = capture()
end
do local __t = tostring(last); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
