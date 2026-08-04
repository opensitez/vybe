-- vybe-test: lua/math_library/math_ceil_for_rounding_up_pages
-- origin: languages/lua/tests/lua/test_math_library.rs

local __w1 = "4"
local __i = 0

local pages = 10
local per = 3
do local __t = tostring(math.ceil(pages / per)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
