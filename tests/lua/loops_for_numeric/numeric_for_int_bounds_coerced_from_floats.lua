-- vybe-test: lua/loops_for_numeric/numeric_for_int_bounds_coerced_from_floats
-- origin: languages/lua/tests/lua/test_loops_for_numeric.rs

local __w1 = "integer,integer,integer,"
local __i = 0

local s = ''
for i = 1.0, 3.0, 1.0 do s = s .. math.type(i) .. ',' end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
