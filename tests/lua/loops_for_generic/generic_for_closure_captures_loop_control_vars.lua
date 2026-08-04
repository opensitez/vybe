-- vybe-test: lua/loops_for_generic/generic_for_closure_captures_loop_control_vars
-- origin: languages/lua/tests/lua/test_loops_for_generic.rs

local __w1 = "1x,2y,3z,"
local __i = 0

local out = ''
for i, v in ipairs({'x', 'y', 'z'}) do
  out = out .. i .. v .. ','
end
do local __t = tostring(out); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
