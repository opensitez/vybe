-- vybe-test: lua/string_patterns_advanced/test_string_gmatch_empty_matches
-- origin: languages/lua/tests/lua/test_string_patterns_advanced.rs

local __w1 = "a, "
local __i = 0

local s='' for w in string.gmatch('a', '.*') do s=s..w..',' end do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
