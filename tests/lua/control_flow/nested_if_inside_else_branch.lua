-- vybe-test: lua/control_flow/nested_if_inside_else_branch
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "mid"
local __i = 0

local x = 5
local result
if x < 3 then
  result = 'low'
else
  if x < 7 then
    result = 'mid'
  else
    result = 'high'
  end
end
do local __t = tostring(result); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
