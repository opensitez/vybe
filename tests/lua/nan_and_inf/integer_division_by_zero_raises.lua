-- vybe-test: lua/nan_and_inf/integer_division_by_zero_raises
-- origin: languages/lua/tests/lua/test_nan_and_inf.rs

local __w1 = "false"
local __i = 0

local ok = pcall(function() return 1 // 0 end)
do local __t = tostring(ok); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
