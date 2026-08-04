-- vybe-test: lua/control_flow/while_with_empty_string_condition_runs_until_break
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "2"
local __i = 0

local c = 0
while "" do
  c = c + 1
  if c == 2 then break end
end
do local __t = tostring(c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
