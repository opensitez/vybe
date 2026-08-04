-- vybe-test: lua/string_rep_rev/test_upper_basic
-- origin: languages/lua/tests/lua/test_string_rep_rev.rs

local __w1 = "ABCD!"
local __i = 0

do local __t = tostring(string.upper('AbCd!')); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
