-- vybe-test: lua/programs/pattern_extract_digits_from_id
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "42"
local __i = 0

do local __t = tostring(string.match("user-42", "%d+")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
