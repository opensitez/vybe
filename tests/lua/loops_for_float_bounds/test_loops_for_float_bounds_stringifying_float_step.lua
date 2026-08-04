-- vybe-test: lua/loops_for_float_bounds/test_loops_for_float_bounds_stringifying_float_step
-- origin: languages/lua/tests/lua/test_loops_for_float_bounds.rs

local __w1 = "1;2;3;"
local __i = 0

local out = ""
for i = 1.0, 3.0, 1.0 do out = out .. tostring(i) .. ';' end
do local __t = tostring(out); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
