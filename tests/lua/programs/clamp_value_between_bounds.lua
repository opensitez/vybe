-- vybe-test: lua/programs/clamp_value_between_bounds
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "10"
local __i = 0

local v = 15
local lo, hi = 0, 10
if v < lo then v = lo elseif v > hi then v = hi end
do local __t = tostring(v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
