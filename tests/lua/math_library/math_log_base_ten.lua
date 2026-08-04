-- vybe-test: lua/math_library/math_log_base_ten
-- origin: languages/lua/tests/lua/test_math_library.rs

local __w1 = "2"
local __i = 0

do local __t = tostring(math.log(100, 10)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
