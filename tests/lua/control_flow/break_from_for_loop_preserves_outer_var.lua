-- vybe-test: lua/control_flow/break_from_for_loop_preserves_outer_var
-- origin: languages/lua/tests/lua/test_control_flow.rs

local __w1 = "preserved"
local __i = 0

local outer = 'preserved'
for i = 1, 5 do
  if i == 3 then break end
end
do local __t = tostring(outer); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
