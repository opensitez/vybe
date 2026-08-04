-- vybe-test: lua/math_fmod_cases/test_math_fmod_cases_rounded
-- origin: languages/lua/tests/lua/test_math_fmod_cases.rs

local __w1 = "true"
local __i = 0

do local __t = tostring(math.fmod(27, 7) == 6); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
